#!/bin/bash

# Check if a filename argument is provided
if [ -z "$1" ]; then
  echo "Usage: $0 filename.docx"
  exit 1
fi

# Assign the first argument to the variable FILENAME
FILENAME="docs/$1"

# Check if the file exists
if [ ! -f "$FILENAME" ]; then
  echo "File not found: $FILENAME"
  exit 1
fi

# Create directory name
DIRNAME="${FILENAME%.*}"

# Check if the directory already exists
if [ -d "$DIRNAME" ]; then
  echo "Directory already exists: $DIRNAME"
  exit 1
fi

# Create the new directory
mkdir "$DIRNAME"
cd "$DIRNAME"

# Unzip the file
unzip ../../"$FILENAME"
cd ../..

# Check if the unzip command was successful
if [ $? -eq 0 ]; then
  echo "File successfully unzipped."
else
  echo "Failed to unzip the file."
  exit 1
fi