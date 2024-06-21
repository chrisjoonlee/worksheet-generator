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

# Get the base directory name without extension
BASE_DIR="${FILENAME%.*}"

# Function to find the next available directory name
get_next_dir_name() {
  local dir_name="$1"
  local i=1
  while [ -d "$dir_name$i" ]; do
    ((i++))
  done
  echo "$dir_name$i"
}

# Check if the directory already exists
if [ -d "$BASE_DIR" ]; then
  BASE_DIR=$(get_next_dir_name "$BASE_DIR")
fi

# Create the new directory
mkdir "$BASE_DIR"
cd "$BASE_DIR"

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