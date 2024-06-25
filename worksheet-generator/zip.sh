#!/bin/bash

# Check if a directory name argument is provided
if [ -z "$1" ]; then
  echo "Usage: $0 directory"
  exit 1
fi

# Assign the first argument to the variable DIRNAME
DIRNAME="docs/$1"

# Check if the directory exists
if [ ! -d "$DIRNAME" ]; then
  echo "Directory not found: $DIRNAME"
  exit 1
fi

cd "$DIRNAME"

# Zip the directory
zip -r ../../"$DIRNAME".docx [Content_Types].xml _rels docProps word

cd ../..

# Check if the zip command was successful
if [ $? -eq 0 ]; then
  echo "Directory successfully zipped into file: $DIRNAME.docx"
else
  echo "Failed to zip the directory."
  exit 1
fi