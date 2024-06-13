#!/bin/bash

# Check if a filename name argument is provided
if [ -z "$1" ]; then
  echo "Usage: $0 filename.docx"
  exit 1
fi

# Assign the first argument to the variable FILENAME
FILENAME="$1"

# Get the base directory name without extension
BASE_DIR="${FILENAME%.*}_new"

# Function to find the next available directory name
get_next_dir_name() {
  local dir_name="$1"
  local i=1
  while [ -d "docs/$dir_name$i" ]; do
    ((i++))
  done
  echo "$dir_name$i"
}

# Check if the directory already exists
if [ -d "docs/$BASE_DIR" ]; then
  BASE_DIR=$(get_next_dir_name "$BASE_DIR")
fi

# ./unzip.sh "$FILENAME"
./unzip.sh "$FILENAME"
dotnet run "$BASE_DIR"
./zip.sh "$BASE_DIR"