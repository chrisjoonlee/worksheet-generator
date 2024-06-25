HOW TO RUN THE WORKSHEET GENERATOR

1.  Place your docx file in the docs folder

2.  Run the following command:
    ./run.sh <file_name>

    Example:
    ./run.sh example.docx

    If you've already run the command, make sure to close the newly
    generated file before running the command again.

3.  The newly generated worksheet will appear in the docs folder with the
    appended suffix "_new"

    Example:
    docs/example_new.docx


HOW TO UNZIP A WORD FILE

1.  Place your docx file in the docs folder

2.  Run the following command:
    ./unzip.sh <file_name>

    Example:
    ./unzip.sh example.docx

3.  The newly unzipped folder will appear in the docs folder with the same
    name as the file without the .docx extension

    Example:
    docs/example_new


HOW TO ZIP A WORD folder

1. Make sure the folder is in the docs folder

2.  Run the following command:
    ./zip.sh <folder_name>

    Example"
    ./zip.sh example

3.  The newly zipped docx file will appear in the docs folder