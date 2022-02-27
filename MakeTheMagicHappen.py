#!/usr/bin/python3

# C. Olson 
# Utility to sort through a folder of PDF files, organize and merge
#
# 2-26.2022
version = "v2022.2.1" 
#       - Adding Code to provide option for a 3 page report split
#		- Replaced different split functions with a unified function PDFSplit3 that takes page split as argument
#		- added execution timing of program to final output
#		- Removed legacy pdf_splitter and PDFSplit2 functions deprecated for PDFSplit3
# 11-13.2021
#version = "v2021.11.0" 
#       - Making needed changes for multi-page class identification
#       - Modified PDF Bytes Reading function   
#		- Removed some UTF8 decode calls that are apparently no longer needed
# 2-17.2021
#version = "v2.1" 
#       - Added Info and Debug Logging
#       - Changed functionality to use the first digit of the source filename to determine order of output report
#       - Added error checking and remediation if a blank PDF page is trailing in a source report
#       - Added error checking to ensure the source files are properly staged in To_Process
#       - Optimized module usage to remove deprecated functions
#
# 5-30-2020
# v2.0 - Consolidated tasks within Main function
#        - Added multi-page split function with keyword ability to accomodate English 2 page doc
#        - cleaned up rename file logic based on class name
#        - Additional commented debug
#        - Made output filename just the student ID instead of student ID - Name

import os
import shutil
import logging
from io import BytesIO
from io import StringIO
import csv
import time
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger


#PDFsplit3 is v3 designed to take the skip step as the second argument.  It provides a single unified routine in code. Old functions have been deprecated and pruned
def PDFsplit3(path, step):
    # creating pdf reader object
    input_pdf = PdfFileReader(open(path, "rb"))
    num_pages = input_pdf.numPages
    input_dir, filename = os.path.split(path)
    filename = os.path.splitext(filename)[0]   
    intervals = range(0, num_pages, step)
    intervals = dict(enumerate(intervals, 1))

    count = 0
    for key, val in intervals.items():
                output_pdf = PdfFileWriter()
                if key == len(intervals):
                    for i in range(val, num_pages):
                        output_pdf.addPage(input_pdf.getPage(i))
                    nums = f'{val + 1}' if step == 1 else f'{val + 1}-{val + step}'
                    naming = '{}_i_{}.pdf'.format(path, i+1)
                    with open(f'{naming}{nums}.pdf', 'wb') as outfile:
                        output_pdf.write(outfile)
                    count += 1
                else:
                    for i in range(val, intervals[key + 1]):
                        output_pdf.addPage(input_pdf.getPage(i))
                    nums = f'{val + 1}' if step == 1 else f'{val + 1}-{val + step}'
                    naming = '{}_i_{}.pdf'.format(path, i+1)
                    with open(f'{naming}{nums}.pdf', 'wb') as outfile:
                        output_pdf.write(outfile)
                    count += 1

    os.rename(path,path+'.processed')
        
#REPLACEMENT PDF TEXT PARSER - 11-2021
def pdf_to_text(path):

    # PDFMiner boilerplate
    rsrcmgr = PDFResourceManager()
    sio = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, sio, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)

    # Extract text
    fp = open(path, 'rb')
    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
    fp.close()

    # Get text from StringIO
    text = sio.getvalue()

    # Cleanup
    device.close()
    sio.close()

    return text


def rename_file(path):
    #This function is designed to scan the text of the PDF file to grab out key information like student name, class name, etc
    #needed for proper file renaming into student folders for merge preparation
    logging.info("Entering Rename File Function")
    text = pdf_to_text(path)
    #DEBUG - Print Raw Text if Needed
    logging.info('Path is: %s', path)
    logging.info('Text is: %s', text)

    preforder = (os.path.basename(path)[0])
    
    name,classname,rest,rest2 = text.split('\n',3)
    logging.info('Name is: %s', name)
    logging.info('Classname is: %s', classname)
    shortclass = classname.split(' ', 1)[0]

    #This catch sequence was used to work around poor class naming - shifting the keywork to the second work in the class string
    if shortclass == "6th":
        shortclass = classname.split(' ', 2)[1]
        logging.info('Shortclass is: %s', shortclass)

    #Make a Copy of the file with the proper naming
    source=os.path.join(os.path.dirname(os.path.abspath(__file__)),'To_Process',filename)

    logging.info('Source is: %s', source)
    newfilename = name+'_'+preforder+'_'+shortclass+'.pdf'
    logging.info('Newfilename is: %s', newfilename)
    destination = os.path.join(os.path.dirname(os.path.abspath(__file__)),'Processed',name,newfilename)
    logging.info('Destination is: %s', destination)

    #Copy existing file to a newly named one
    os.makedirs(os.path.dirname(destination), exist_ok=True)
    shutil.copy(source,destination)
    os.rename(source,source+'.processed')

def merge_files(dirName):
    #Sort the files alpabetically so they print in the right order
    sfileList = sorted(fileList, reverse = False)
    #Loop through folders and merge
    merger = PdfFileMerger(strict=False)
    logging.info('merge_files - sFileList is: %s', sfileList)

    for fname in sfileList:
        merger.append(PdfFileReader(open(os.path.join(dirName, fname),'rb')))
        logging.info('merge_files - fname is: %s', fname)

    merger.write(str(dirName)+".pdf")
    logging.info("wrote file")

def add_student_id(path):

    #Set some Error Checking Values
    n=0
    m=0
    
    #open and store the csv file
    with open('StudentIDs.csv', 'r') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',')

        for row in csvreader:
            name = row[1] + '.pdf'
            #This would output studentID - name.pdf
            #new = row[0] + ' - ' + row[1] + '.pdf'
            #This will just output studentID.pdf
            new = row[0] + '.pdf'
            logging.info('add_student_id - Name is: %s', name)
            logging.info('add_student_id - New is: %s', new)
            fullpath=os.path.join(path, name)
            fullnew=os.path.join(path, new)
            logging.info('add_student_id - FullPath is: %s', fullpath)
            if os.path.exists(fullpath):
                os.rename(fullpath, fullnew)
                logging.info('*** Match Detected for %s', name)
                n=n+1
            else:
                logging.warning('*** Match NOT Detected for %s', name)
                m=m+1
    
    if m==0:
        print("\n** "+ str(n) +" Student Records Successfully Processed **\n")
    else:
        print("** REVIEW REQUIRED **\n"+ str(m) +" Student Records were not Processed due to name mismatch...\nCheck name in PDF against CSV File")

if __name__ == '__main__':
    startTime = time.time()
    ########### Configuration Section - Please review before each Trimester #######################################################
    #Split the PDF File using the below keywords to identify where we need to split 2 pages instead of 1
    #This needs to be updated each Trimester depending on the way the classes are output - Any File that is 2 page per student instead of one
    #Needs to be identified.  Code will use filename substring search.
    #keywords will split for a 2 page report, keywords3 will split for a 3 page report **This needs to be cleaned up
    keywords2 = ['History', 'Science', 'Chorus']
    keywords3 = ['Math', 'Habits']
    #Set desired logging level.  CRITICAL, WARNING, INFO, DEBUG 
    loglevel = 'WARNING'

    #End Configuration Area ############################################################################################

    loglevel = os.environ.get('LOGLEVEL', loglevel).upper()
    logging.basicConfig(level=loglevel)

    #Start the Main Program
    print ("****** SBG Report Generator " +version+ " Beginning ... ")
    
    #Set the target staging location for source files to the To_Process Subfolder
    directory=os.path.join(os.path.dirname(os.path.abspath(__file__)),'To_Process')

    #Make sure the source files are staged
    if len(next(os.walk(directory))[2]) <= 1:
        logging.critical("\nNo Files Detected in To_Process Folder to process, Exiting...\n\n")
        raise SystemExit
    else:
        print ("****** Go get a cup of Coffee!! ******\n")
        time.sleep(1)

    print ("\n* Processing Source Class Files *")
    print ("\n*******  Will process 2-page split for the following classes")
    print (*keywords2, sep= ", ")
    print ("\n*******  Will process 3-page split for the following classes")
    print (*keywords3, sep= ", ")
      
    print ("\n")
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            print ('*** Splitting File **'+ filename)
            path = os.path.join(directory,filename)

            #if filename.contains(keywords) == True:   - This keys off the configured classes above that need multi-page splitting
            if any(substring in filename for substring in keywords2) == True:
                #Call Multi Page Split Function- This will print every other page in a 2 page doc
                logging.info("***Multi Page PDF Detected***")
                PDFsplit3(path,2)
            elif any(substring in filename for substring in keywords3) == True:
                #Call Multi Page Split Function- This will print every other page in a 3 page doc
                logging.info("***3 Page PDF Detected***")
                PDFsplit3(path,3)
            else:
                #Call Single Page Split Function
                logging.info("***Single Page PDF Detected***")
                #pdf_splitter(path)
                PDFsplit3(path,1)
        else:
            continue

    time.sleep(1)
    #Rename the PDF Files
    #Update with each filename using Loop within "Processed" Folder
    print ("\n* Identifiying Student Records *")
    time.sleep(1)
    directory=os.path.join(os.path.dirname(os.path.abspath(__file__)),'To_Process')
    logging.info("Directory is "+ directory)
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            logging.info("Renaming File"+ filename)
            rename_file(os.path.join(directory, filename))
            print('.', end='', flush=True)
        else:
            logging.info("File was not a PDF"+ filename)
            continue
    print ("\n")

    #Merge the PDF Files by Student
    #Iterate through each created directory and call PDF Print Action, naming the merged file by the folder name and placing in the root Directory
    print ("* Merging Student Records into consolidated Report *")
    rootDir=os.path.join(os.path.dirname(os.path.abspath(__file__)),'Processed')
    for dirName,subDir, fileList in os.walk(rootDir, topdown=False):
        if dirName != rootDir:    #Exclude the root folder
            logging.info('dirName is %s', dirName)
            merge_files(dirName)
            print('.', end='', flush=True)
    print ("\n* Student Record Consolidation Complete *")

    #Call Add_student_id method to scan each folder, compare against student ID CSV File, and rename the files properly
    print ("* Appending Student ID to Final Report Filename *")
    add_student_id(rootDir)
    print ("* Renaming Operation Completed *")

     #MicDrop
    executionTime = (time.time() - startTime)
print ("****** You just saved 8 hours! The Magic Just Happened in " + str(executionTime) +" seconds!")
