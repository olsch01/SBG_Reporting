# SBG_Reporting
Utility designed to turn Full Class Standards Based Grading reports from Powerschool's SeniorSystem application into Student Centric reports by splitting, processing, organizing and combining PDF Reports to finalized Student Reports for publishing.

Use the following Steps to Stage and Execute SBG_Reporting
1. Clone this repository
2. Install python3 requirements as provided in requirements.txt with the following command:
	pip3 install requirements.txt
3. Place All source class reports into the "To_Process" folder
4. Review the source documents for the following:
	1) If specific ordering is desired, use a leading number to set the order precedent:
		- IE, if you like English first and Algebra Second, rename the source files to:
			1_English.pdf
			2_Algebra.pdf
		Otherwise, Alphabetical order will be used.
	2) Review each file to determine if each students record is only one page, or if it is two pages.  If it is two pages, you'll need to add a substring of the filename into the "keywords" variable list on line 187.  Use the existing example as a guide.   
	* If a keyword is not provided, the default logic is to split each page
5. To facilitate naming the final output files as the student ID, populate the StudentIDs.csv file with a comma delineated formate of StudentID,FormalName (as shown in the Powerschool Report)
		NOTE: It is important that the names match identically, otherwise the program will alert you that it cannot find a match, and the final PDF will not be renamed to a studentID and will be kept with the detected student name only.
6. Once everything is staged, execute ./MakeTheMagicHappen.py

The program will process the source files into various output files in the "Processed" folder.  *The final step will create a bunch of <studentid>.pdf files which are the final output**.  No intermediary files are removed by default in the event you need to review.
