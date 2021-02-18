#!/usr/bin/python3

# C. Olson 5-30-2020
# Utility to sort through a folder of PDF files, organize and merge
# v2.0 - Consolidated tasks within Main function
#        - Added multi-page split function with keyword ability to accomodate English 2 page doc
#        - cleaned up rename file logic based on class name
#        - Additional commented debug
#        - Made output filename just the student ID instead of student ID - Name

import os
import shutil
import logging
from io import BytesIO
#import glob
#import sys
import csv
#import re
import time
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger


#Split up a PDF File passed in as parameter page by page
def pdf_splitter(path):
#    fname = os.path.splitext(os.path.basename(path))[0]
    fname = path
    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))

        output_filename = '{}_page_{}.pdf'.format(fname, page+1)

        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)

        logging.info('Created: {}'.format(output_filename))
        
        #Function to work around a blank PDF page - this came up with a blank page that
        #when split, would cause an exception
        statinfo = os.stat(output_filename)
        if statinfo.st_size <= 1024:
            logging.info('******* Blank Page Detected, deleting*')
            os.remove(output_filename)
        
    os.rename(path,path+'.processed')

def PDFsplit2(path):

    # creating pdf reader object
    inputFile = PdfFileReader(open(path, "rb"))

    for i in range(inputFile.numPages // 2):
        output = PdfFileWriter()
        output.addPage(inputFile.getPage(i * 2))

        if i * 2 + 1 <  inputFile.numPages:
            output.addPage(inputFile.getPage(i * 2 + 1))

        output_filename = '{}_i_{}.pdf'.format(path, i+1)

        with open(output_filename, 'wb') as out:
            output.write(out)

        #Function to work around a blank PDF page - this came up with a blank page that
        #when split, would cause an exception
        statinfo = os.stat(output_filename)
        if statinfo.st_size <= 1024:
            logging.info('******* Blank Page Detected, deleting*')
            os.remove(output_filename)

    os.rename(path,path+'.processed')

#Convert PDF to Text to parse
def pdf_to_text(path):
    manager = PDFResourceManager()
    retstr = BytesIO()
    layout = LAParams(all_texts=True)
    device = TextConverter(manager, retstr, laparams=layout)
    filepath = open(path, 'rb')
    interpreter = PDFPageInterpreter(manager, device)

    for page in PDFPage.get_pages(filepath, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    filepath.close()
    device.close()
    retstr.close()
    return text

def rename_file(path):
    text = pdf_to_text(path)
    #DEBUG - Print Raw Text if Needed
    logging.info('Path is: %s', path)
    logging.info('Text is: %s', text)

    name,classname,rest,rest2 = text.decode('utf8').split('\n',3)
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
    newfilename = name+'_'+shortclass+'.pdf'
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
    merger = PdfFileMerger()

    for fname in sfileList:
        merger.append(PdfFileReader(open(os.path.join(dirName, fname),'rb')))
        logging.info('fname is: %s', fname)
        merger.write(str(dirName)+".pdf")

def add_student_id(path):

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
            else:
                logging.warning('*** Match NOT Detected for %s', name)

if __name__ == '__main__':
    #Configuration Section - Please review before each Trimester #######################################################
    #Split the PDF File using the below keywords to identify where we need to split 2 pages instead of 1
    #This needs to be updated each Trimester depending on the way the classes are output
    keywords = ['Math2', 'Arts_Chorus','WL_French']


    #End Configuration Area ############################################################################################

    loglevel = 'WARNING'
    #loglevel = 'INFO'

    loglevel = os.environ.get('LOGLEVEL', loglevel).upper()
    logging.basicConfig(level=loglevel)

    #Start the Main Program

    print ("****** Go get a cup of Coffee!! ******")
    time.sleep(2)

    directory=os.path.join(os.path.dirname(os.path.abspath(__file__)),'To_Process')

    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            print ('*** Splitting File **'+ filename)
            path = os.path.join(directory,filename)

            #if filename.contains(keywords) == True:   - This keys off the configured classes above that need multi-page splitting
            if any(substring in filename for substring in keywords) == True:
                #Call Multi Page Split Function- This will print every other page in a 2 page doc
                logging.info("***Multi Page PDF Detected***")
                PDFsplit2(path)
            else:
                #Call Single Page Split Function
                logging.info("***Single Page PDF Detected***")
                pdf_splitter(path)
        else:
            continue

    print ("\n* Class File Splitting Complete *")
    time.sleep(2)
    #Rename the PDF Files
    #Update with each filename using Loop within "Processed" Folder
    print ("* Identifiying Student Records *")
    time.sleep(3)
    directory=os.path.join(os.path.dirname(os.path.abspath(__file__)),'To_Process')
    logging.info("Directory is "+ directory)
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            rename_file(os.path.join(directory, filename))
            print('.', end='', flush=True)
        else:
            continue
    print ("\n* Student Record Identification Complete *")
    time.sleep(2)

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
    time.sleep(2)
    print ("* Appending Student ID to Final Report Filename *")
    #Call Add_student_id method to scan each folder, compare against student ID CSV File, and rename the files properly
    add_student_id(rootDir)
    print ("* Renaming Operation Completed *")

     #MicDrop
print ("****** The Magic Just Happened ******")
