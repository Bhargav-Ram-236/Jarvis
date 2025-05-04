import os
import random
import time
import subprocess

def play_random_file(folder_path):
    """Play a random media file from the specified folder without opening a window."""
    if not os.path.exists(folder_path):
        print("The provided folder path does not exist.")
        return
    
    # List all files in the folder (any file type)
    files = os.listdir(folder_path)
    
    if not files:
        print("No files found in the folder.")
        return
    
    # Randomly pick a file
    random_file = random.choice(files)
    file_path = os.path.join(folder_path, random_file)

    # Play the file using available media players
    try:
        # Try using Windows Media Player
        subprocess.Popen([r"C:\Program Files\Windows Media Player\wmplayer.exe", file_path], shell=True)
    except FileNotFoundError:
        try:
            # Try using VLC (if installed)
            subprocess.Popen([r"C:\Program Files\VideoLAN\VLC\vlc.exe", file_path], shell=True)
        except FileNotFoundError:
            print("No supported media player found.")

    # Give some time before proceeding (Optional)
    time.sleep(2)

# Directly execute the action
if __name__ == "__main__":
    play_random_file(r"C:\My Certificates\scolding")  # Change this to your folder path
