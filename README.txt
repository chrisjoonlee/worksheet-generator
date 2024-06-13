HOW TO RUN THE WORKSHEET GENERATOR

1.  Place your docx file in the docs folder

2.  Run the following command:
    ./run.sh <file_name>

    For example:
    ./run.sh example.docx

3.  The newly generated worksheet will appear in the docs folder with the
    appended suffix "_new#"

    For example:
    docs/example_new1.docx


HOW TO CREATE A NEW WORD DOC

1.  Copy the starter:
    ./newDoc.sh <new_file_name>

    For example:
    ./newDoc.sh example.docx

    If something is wrong with the starter folder, simply create a new one:
    - Create an empty Word docx file called "starter.docx"
    - Run the following command:
      ./unzip.sh starter


HOW TO EDIT A WORD DOC

1.  Unzip the word doc:
    ./unzip.sh <file_name>

    For example:
    ./unzip.sh example.docx

2.  A newly generated folder will appear in the docs folder with the appended
    suffix "_new#"

    For example:
    docs/example_new1

3.  Edit and save the XML files in an editor.
    Do not open the folder in Finder.

4.  Zip the word doc:
    ./zip.sh <folder_name>

    For example:
    ./zip.sh example_new1