#!/bin/bash

# Check if the correct number of arguments is provided
if [ "$#" -ne 1 ]; then
    echo "Usage: $0 new_folder_name"
    exit 1
fi

SOURCE_DIR="starter"
TARGET_DIR=$1

# Check if the source directory exists
if [ ! -d "$SOURCE_DIR" ]; then
    echo "Error: Source directory '$SOURCE_DIR' does not exist."
    exit 1
fi

# Copy the directory and its contents
cp -r "$SOURCE_DIR" "$TARGET_DIR"

# Check if the copy was successful
if [ $? -eq 0 ]; then
    echo "Successfully created folder '$TARGET_DIR'."
else
    echo "Error: Failed to create folder '$TARGET_DIR'."
    exit 1
fi