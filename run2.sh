# Check if a filename name argument is provided
if [ -z "$1" ]; then
  echo "Usage: $0 filename.docx"
  exit 1
fi

# Assign file names
FILENAME="$1"
BASE_FILENAME="${FILENAME%.*}_new"

# Function to find the next available directory name
get_next_dir_name() {
  local file_name="$1"
  local i=1
  while [ -d "docs/$file_name$i" ]; do
    ((i++))
  done
  echo "$file_name$i"
}

# Check if the directory already exists
if [ -d "docs/$BASE_FILENAME" ]; then
  BASE_FILENAME=$(get_next_dir_name "$BASE_FILENAME")
fi

echo dotnet run "$FILENAME" "$BASE_FILENAME.docx"
dotnet run "$FILENAME" "$BASE_FILENAME.docx"