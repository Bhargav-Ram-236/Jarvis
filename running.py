import subprocess
import os

# Absolute path to your React app folder
react_app_path = r"C:\My Certificates\Jarvis\todo"  # change this to your folder path

# Run `npm start` inside the React app folder
process = subprocess.Popen(
    ["npm", "start"],
    cwd=react_app_path,
    shell=True
)
