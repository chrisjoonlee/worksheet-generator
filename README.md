### Worksheet Generator

This proof-of-concept program converts plain-text Microsoft Word documents into formatted worksheets. It automates the styling process, shuffles choices in multiple-choice questions, and generates answer keys. This tool is designed for teachers who want to simplify the creation of their worksheets.

## How to run the program

0.  Clone and enter the repo
1.  Place your docx file in the `docs` folder
2.  Run `./run.sh <file_name>`

    Example: `./run.sh example.docx`
    
    If you've already run the command, make sure to close the newly generated file before running the command again.

3.  The newly generated worksheet will appear in the docs folder with the appended suffix "_new"
    Example: `docs/example_new.docx`


## How to unzip a Word file

0.  Clone and enter the repo
1.  Place your docx file in the docs folder
2.  Run `./unzip.sh <file_name>`
   
    Example: `./unzip.sh example.docx`

4.  The newly unzipped folder will appear in the docs folder with the same
    name as the file without the .docx extension
    Example: `docs/example_new`


## How to zip a Word folder

0.  Clone and enter the repo
1.  Make sure the folder is in the docs folder
2.  Run `./zip.sh <folder_name>`

    Example: `./zip.sh example`

4.  The newly zipped docx file will appear in the docs folder
