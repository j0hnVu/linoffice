#!/bin/bash

echo "This script will now delete the LinOffice podman container and all its data, using the command 'podman rm -f LinOffice && podman volume rm linoffice_data'"

read -p "Are you sure you want to proceed? (y/N): " response
response=$(echo "$response" | tr '[:upper:]' '[:lower:]')

if [[ "$response" == "y" || "$response" == "yes" ]]; then
    echo "Deleting LinOffice container and data..."
    podman rm -f LinOffice && podman volume rm linoffice_data
    if [ $? -eq 0 ]; then
        echo "Successfully deleted LinOffice container and data."
    else
        echo "Error: Failed to delete LinOffice container or data."
        exit 1
    fi
else
    echo "Operation aborted by user."
    exit 0
fi
