import os
import subprocess
import sys

# Predefined common directories
COMMON_PATHS = {
    "Desktop": r"C:\Users\bharg\OneDrive\Desktop",
    "Downloads": r"C:\Users\bharg\Downloads",
    "Document": r"C:\Users\bharg\OneDrive\Documents",
    "Pictures": r"C:\Users\bharg\OneDrive\Pictures"
}

def find_and_open_folder(folder_name):
    """Search in predefined directories and open the folder if found."""

    if not folder_name:
        print("Please specify a folder name.")
        return
    
    folder_name_lower = folder_name.lower()
    
    print(f"Searching for: {folder_name}")

    # Check if folder_name directly matches one of the common paths
    for key, path in COMMON_PATHS.items():
        if folder_name_lower == key.lower():  # Direct match with predefined directories
            print(f"FOUND: Opening {key} ({path})")
            subprocess.Popen(f'explorer "{path}"', shell=True)
            return

    possible_matches = []

    # Check if the folder exists inside predefined directories
    for name, path in COMMON_PATHS.items():
        folder_path = os.path.join(path, folder_name)
        print(f"Checking: {folder_path}")  # Debugging: Print checked paths

        if os.path.exists(folder_path):
            print(f"FOUND in {name}: {folder_path}")  # Debugging: Found location
            subprocess.Popen(f'explorer "{folder_path}"', shell=True)
            return

    # If not found, search in common directories
    for base_dir in COMMON_PATHS.values():
        for root, dirs, _ in os.walk(base_dir):
            for d in dirs:
                if folder_name_lower in d.lower():
                    match_path = os.path.join(root, d)
                    possible_matches.append(match_path)

    # Open if exact match found, otherwise suggest alternatives
    if possible_matches:
        print("Did you mean one of these?")
        for match in possible_matches:
            print(match) 
    else:
        print(f"Folder '{folder_name}' not found.")

# Get folder name from command-line argument
if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_name = sys.argv[1]  # Takes only the folder name
        find_and_open_folder(folder_name)
    else:
        print("No folder name provided. Usage: python folder.py <folder-name>")
