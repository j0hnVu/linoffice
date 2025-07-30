#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPT_NAME="$(basename "$0")"
echo "Executing uninstall script in $SCRIPT_DIR"

# Check if setup.sh exists and extract USER_APPLICATIONS_DIR and APPDATA_PATH
if [[ ! -f "$SCRIPT_DIR/setup.sh" ]]; then
  echo "Warning: setup.sh not found in the current directory."
else
  eval $(grep -E '^\s*USER_APPLICATIONS_DIR=' $SCRIPT_DIR/setup.sh)
  eval $(grep -E '^\s*APPDATA_PATH=' $SCRIPT_DIR/setup.sh)

  if [[ -z "$USER_APPLICATIONS_DIR" ]]; then
    echo "Warning: USER_APPLICATIONS_DIR not found in setup.sh."
  fi

  if [[ -z "$APPDATA_PATH" ]]; then
    echo "Warning: APPDATA_PATH not found in setup.sh."
  fi
fi

# Check if APPDATA_PATH exists and delete if it does
if [[ -d "$APPDATA_PATH" ]]; then
  read -p "Do you want to delete the app data in $APPDATA_PATH? (y/n): " confirm
  if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
    rm -r "$APPDATA_PATH"
    echo "Deleted directory: $APPDATA_PATH"
  else
    echo "Deletion of $APPDATA_PATH aborted."
  fi
else
  echo "Warning: Directory $APPDATA_PATH does not exist."
fi

# Find .desktop files containing linoffice.sh in Exec= line
if [[ -n "$USER_APPLICATIONS_DIR" ]]; then
  DESKTOP_FILES=$(find "$USER_APPLICATIONS_DIR" -type f -name "*.desktop" -exec grep -l "Exec=.*linoffice.sh" {} \;)
  if [[ -n "$DESKTOP_FILES" ]]; then
    echo "The following .desktop files will be deleted:"
    echo "$DESKTOP_FILES"
    read -p "Do you want to proceed with deletion? (y/n): " confirm
    if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
      while IFS= read -r file; do
        rm -f "$file" && echo "Deleted: $file"
      done <<< "$DESKTOP_FILES"
    else
      echo "Deletion of .desktop files aborted."
    fi
  else
    echo "No .desktop files containing linoffice.sh found."
  fi
fi

# Ask to delete the Windows container and its data
read -p "Do you want to delete the Windows container and all its data as well? (y/n): " confirm
if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
  if ! command -v podman &> /dev/null; then
    echo "Error: Podman is not installed or not accessible."
  else
    # Stop and remove the LinOffice container
    if ! podman stop LinOffice &> /dev/null; then
      echo "Error: Could not stop the LinOffice container. It may not exist."
    else
      echo "Stopped LinOffice container."
    fi

    if ! podman rm -f LinOffice &> /dev/null; then
      echo "Error: Could not delete the LinOffice container."
    else
      echo "Deleted LinOffice container."
    fi

    # Remove the linoffice_data volume
    if ! podman volume rm linoffice_data &> /dev/null; then
      echo "Error: Could not delete the linoffice_data volume."
    else
      echo "Deleted linoffice_data volume."
    fi
  fi
else
  echo "Windows container and data deletion aborted."
fi

# Find all files and folders in the same directory as uninstall.sh (excluding itself)

FILES_TO_DELETE=$(find "$SCRIPT_DIR" -maxdepth 1 -not -name "$SCRIPT_NAME")
if [[ -n "$FILES_TO_DELETE" ]]; then
  echo "The following files and folders will be deleted recursively:"
  echo "$FILES_TO_DELETE"
  read -p "Do you want to proceed with deletion? (y/n): " confirm
  if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
    find "$SCRIPT_DIR" -maxdepth 1 -not -name "$SCRIPT_NAME" -exec rm -rf {} \;
    if [[ -f "$SCRIPT_DIR/setup.sh" ]]; then
      rm -f "$SCRIPT_DIR/setup.sh"
    fi
    echo "Files and folders deleted."
    # Delete the uninstall.sh script itself
    echo "Deleting the uninstall script itself."
    rm -f "$0"
    echo "Uninstall script deleted."
  else
    echo "Deletion of files and folders aborted."
  fi
else
  echo "No files or folders to delete in $SCRIPT_DIR."
fi

exit 0
