import speech_recognition as sr
import subprocess
import os
import subprocess
import pyttsx3
import speech_recognition as sr
import random
import pyautogui
import datetime
import webbrowser
import pyjokes
import pygetwindow as gw 
import time
import schedule
from datetime import datetime
import threading
import sounddevice as sd
import vosk
import json
import requests
import google.generativeai as genai

# Initialize TTS
engine = pyttsx3.init('sapi5')
engine.setProperty('rate', 180)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Use female voice

# Memory stacks
opened_tabs_stack = []
closed_tabs_stack = []
last_actions_stack = []
last_opened_app = ""

# Speak function
def speak(text):
    print(f"Jarvis: {text}")
    engine.say(text)
    engine.runAndWait()
    time.sleep(0.1)

#Listen for any command that includes "jarvis"
jarvis_active_until = 0
def listen():
    global jarvis_active_until
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎙️ Listening")
        recognizer.adjust_for_ambient_noise(source)
        while True:
            try:
                audio = recognizer.listen(source, phrase_time_limit=5)
                command = recognizer.recognize_google(audio).lower()
                print(f"You said: {command}")
                 
                current_time = time.time()

                if "jarvis" in command:
                    jarvis_active_until = current_time + 60 
                    return command.replace("jarvis", "").strip()

                elif current_time < jarvis_active_until:
                    return command.strip()

                else:
                    print("⏳ Waiting for wake word 'Jarvis'...")
                    continue

            except sr.UnknownValueError:
                continue
            except sr.RequestError:
                speak("Please check your internet connection.")
                continue


# Replace with your actual API key


#Search in windows
def search_in_windows(query):
    pyautogui.press('win')  # Press the Windows key
    time.sleep(0.5)          # Short delay to allow search bar to open
    pyautogui.write(query)   # Type the search term
    time.sleep(0.3)
    pyautogui.press('enter') 
def is_notepad_active():
    win = gw.getActiveWindow()
    return win and "Notepad" in win.title

def ask_gemini(prompt):
    model = genai.GenerativeModel(model_name="gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text

def open_application(app_name):
    global last_opened_app

    # Apps to open via Windows search
    search_based_apps = [
        "whatsapp", "spotify", "telegram", "instagram", "youtube", "zoom",
        "discord", "obs", "photos", "vlc", "bluetooth", "onenote"
    ]

    # Apps to open directly
    app_functions = {
        "notepad": "notepad",
        "calculator": "calc",
        "paint": "mspaint",
        "file explorer": "explorer",
        "settings": "start ms-settings:",
        "camera": "start microsoft.windows.camera:",
        "edge": "msedge",
        "cmd": "C:\\Windows\\System32\\cmd.exe",
        "task manager": "taskmgr",
        "control panel": "control",
        "wordpad": "write",
        "powershell": "powershell",
        "microsoft store": "start ms-windows-store:",
        "snipping tool": "snippingtool",
        "media player": "wmplayer",
        "clock": "start ms-clock:",
        "mail": "start outlookmail:",
        "maps": "start bingmaps:",
        "weather": "start bingweather:",
        "google chrome": r'"C:\Program Files\Google\Chrome\Application\chrome.exe"',
        "visual studio": "code"
    }

    # First, check direct launch apps
    for key, command in app_functions.items():
        if key in app_name:
            speak(f"Opening {key.capitalize()}")
            subprocess.Popen(command, shell=True)
            opened_tabs_stack.append(key)
            last_opened_app = key
            last_actions_stack.append(lambda: open_application(key))
            return

    # Then, check apps to launch via search bar
    for keyword in search_based_apps:
        if keyword in app_name:
            speak(f"Searching and opening {keyword.capitalize()} via Windows search")
            search_in_windows(keyword)
            opened_tabs_stack.append(keyword)
            last_opened_app = keyword
            last_actions_stack.append(lambda: open_application(keyword))
            return

    if "open folder" in app_name:
        folder_name = app_name.replace("open folder", "").strip()
        try:
            subprocess.Popen(["python", "folder.py", folder_name])
            speak(f"Opening folder: {folder_name}")
            last_actions_stack.append(lambda: subprocess.Popen(["python", "folder.py", folder_name]))
        except:
            speak(f"Error opening folder: {folder_name}")
        return

    speak(f"Sorry, I couldn't find {app_name}.")


# Close apps
def close_application(app_name):
    global opened_tabs_stack, closed_tabs_stack

    close_commands = {
        "notepad": "taskkill /F /IM notepad.exe",
        "calculator": "taskkill /F /IM ApplicationFrameHost.exe",
        "paint": "taskkill /F /IM mspaint.exe",
        "file explorer": "taskkill /F /IM explorer.exe",
        "cmd": "taskkill /F /IM cmd.exe",
        "edge": "taskkill /F /IM msedge.exe",
        "task manager": "taskkill /F /IM Taskmgr.exe",
        "control panel": "taskkill /F /IM control.exe",
        "wordpad": "taskkill /F /IM wordpad.exe",
        "powershell": "taskkill /F /IM powershell.exe",
        "media player": "taskkill /F /IM wmplayer.exe",
        "google chrome": "taskkill /F /IM chrome.exe",
        "visual studio": "taskkill /F /IM Code.exe",
        "camera": "taskkill /F /IM WindowsCamera.exe",
        "snipping tool": "taskkill /F /IM SnippingTool.exe",
        "whatsapp":"taskkill /F /IM WhatsApp.exe"
    }

    if "close tab" in app_name or "close window" in app_name or "close" in app_name:
        if opened_tabs_stack:
            last_opened = opened_tabs_stack.pop()
            if isinstance(last_opened,str):

                process_name = last_opened.lower()
                command = close_commands.get(process_name, "")
                if command:
                    speak(f"Closing {last_opened.capitalize()}")
                    subprocess.run(command, shell=True)
                    closed_tabs_stack.append(last_opened)
                    last_actions_stack.append(lambda: close_application(f"close {last_opened}"))
                else:
                    speak(f"Sorry , I can't close{last_opened} right now.")
            else:
                speak(f"Sorry, I can't close {last_opened} right now.")
        else:
            speak("No application is currently opened.")

def open_closed_tab():
    """Reopen the last closed tab."""
    global last_opened_app

    if closed_tabs_stack:
        last_closed_tab = closed_tabs_stack.pop()
        speak(f"Reopening {last_closed_tab.capitalize()}")
        open_application(last_closed_tab)
        last_actions_stack.append(lambda: open_closed_tab())  # ✨
    else:
        speak("No closed application found to reopen.")

def close_active_window():
    """Close the active window."""
    speak("Closing the active window.")
    pyautogui.hotkey('alt', 'f4')
    last_actions_stack.append(lambda: close_active_window())  # ✨

def exit_jarvis():
    """Exit the assistant."""
    speak("Good bye Ram! Have a wonderful day.")
    exit()

def respond_to_gratitude():
    """Respond to gratitude from the user."""
    gratitude_responses = [
        "You're welcome, Ram! 😊",
        "Always happy to help!",
        "No problem at all!",
        "My pleasure!",
        "Glad I could assist!",
        "Anytime, Ram!",
        "You're the best! 🚀",
        "No need to thank me, just doing my duty!",
    ]
    speak(random.choice(gratitude_responses))

def repeat_last_action():
    if last_actions_stack:
        action = last_actions_stack[-1]  # Get last action (do not pop)
        action_type = action[0]
        window = action[1]

        if action_type == 'minimize':
            window.minimize()
            speak("Repeating minimize window.")
        elif action_type == 'maximize':
            window.maximize()
            speak("Repeating maximize window.")
        elif action_type == 'zoom':
            window.resizeTo(window.width + 200, window.height + 200)
            speak("Repeating zoom window.")
        elif action_type == 'zoom_back':
            window.resizeTo(max(200, window.width - 200), max(200, window.height - 200))
            speak("Repeating zoom back window.")
        elif action_type == 'move_left':
            window.moveTo(window.left - 100, window.top)
            speak("Repeating move left.")
        elif action_type == 'move_right':
            window.moveTo(window.left + 100, window.top)
            speak("Repeating move right.")
        else:
            speak("Cannot repeat that action.")
    else:
        speak("No action to repeat.")

def greet_user():
    """Greet the user when Jarvis starts."""
    speak("Hello Ram! I am Jarvis, your personal assistant. How can I assist you today?")

# ============================
# New Features
# ============================

def search_google(query):
    """Search a query on Google."""
    query=query.replace("search for","").strip()
    speak(f"Searching Google for {query}")
    webbrowser.open(f"https://www.google.com/search?q={query}")
    last_actions_stack.append(lambda: search_google(query))  # ✨
    opened_tabs_stack.append("edge")

def play_on_youtube(query):
    """Play a video on YouTube."""
    speak(f"Playing {query} on YouTube")
    webbrowser.open(f"https://www.youtube.com/results?search_query={query}")
    last_actions_stack.append(lambda: play_on_youtube(query))  # ✨

def tell_time():
    """Tell the current time."""
    time_now = datetime.now().strftime("%I:%M %p")
    speak(f"The time is {time_now}")
    last_actions_stack.append(lambda: tell_time())  # ✨

def tell_date():
    """Tell today's date."""
    today = datetime.now().strftime("%B %d, %Y")
    speak(f"Today's date is {today}")
    last_actions_stack.append(lambda: tell_date())  # ✨

def take_screenshot():
    """Take a screenshot and save it."""
    speak("Taking a screenshot.")
    screenshot = pyautogui.screenshot()
    screenshot.save("screenshot.png")
    speak("Screenshot saved.")
    last_actions_stack.append(lambda: take_screenshot())  # ✨

def tell_joke():
    """Tell a joke."""
    joke = pyjokes.get_joke()
    speak(joke)
    last_actions_stack.append(lambda: tell_joke())  # ✨


def minimize_window():
    """Minimize the active window."""
    current_window = gw.getActiveWindow()
    if current_window:
        current_window.minimize()
        speak("Minimizing the active window.")
        last_actions_stack.append(('minimize', current_window))   # ✨

def maximize_window():
    """Maximize the active window."""
    current_window = gw.getActiveWindow()
    if current_window:
        current_window.maximize()
        speak("Maximizing the active window.")
        last_actions_stack.append(('maximize', current_window))  # ✨

def zoom_window():
    window = gw.getActiveWindow()
    if window:
        # Save old size and position
        old_box = (window.left, window.top, window.width, window.height)
        window.resizeTo(window.width + 200, window.height + 200)
        speak("Zooming the window.")
        last_actions_stack.append(('zoom', window, old_box))

def zoom_back_window():
    window = gw.getActiveWindow()
    if window:
        old_box = (window.left, window.top, window.width, window.height)
        window.resizeTo(max(200, window.width - 200), max(200, window.height - 200))
        speak("Zooming back the window.")
        last_actions_stack.append(('zoom_back', window, old_box))

def move_window_left():
    window = gw.getActiveWindow()
    if window:
        old_box = (window.left, window.top, window.width, window.height)
        window.moveTo(window.left - 100, window.top)
        speak("Moving window to the left.")
        last_actions_stack.append(('move_left', window, old_box))

def move_window_right():
    window = gw.getActiveWindow()
    if window:
        old_box = (window.left, window.top, window.width, window.height)
        window.moveTo(window.left + 100, window.top)
        speak("Moving window to the right.")
        last_actions_stack.append(('move_right', window, old_box))

def list_open_windows():
    windows = gw.getAllTitles()  # Get all window titles
    return windows
def schedule_task():
    """Function to schedule a task."""
    # Schedule an example task: this will run every minute as an example.
    schedule.every(1).minute.do(task_to_schedule)

def task_to_schedule():
    """The task you want to run on schedule."""
    speak("Scheduled task is now running!")

def listen_for_schedule():
    """Listen for 'schedule' to activate task scheduling."""
    while True:
        command = listen()  # Use your existing listen() function
        if "schedule" in command:
            speak("Scheduling the task.")
            schedule_task()
            break


# Function to switch to a specific window
def switch_to_window(window_title):
    windows = gw.getWindowsWithTitle(window_title)  # Find windows matching the title
    if windows:
        window = windows[0]  # Assuming the first match
        window.activate()  # Bring the window to the front
        print(f"Switched to window: {window_title}")
    else:
        print(f"No window found with title: {window_title}")

# Function to switch to the last used app (this would work if you have a list of recently used windows)
def switch_to_last_window():
    windows = gw.getAllWindows()
    if windows:
        last_window = windows[-1]  # Pop the last window
        last_window.activate()  # Bring the last window to the front
        print(f"Switched to the last window: {last_window.title}")
    else:
        print("No windows to switch to.")

# Function to switch using "Alt + Tab" to cycle through apps
def alt_tab_switch():
    pyautogui.hotkey('alt', 'tab')  # Simulate pressing Alt + Tab to switch apps
    time.sleep(1)  # Wait for the tab to switch


def abort_last_action():
    if last_actions_stack:
        action = last_actions_stack.pop()
        action_type = action[0]
        window = action[1]

        if action_type in ['minimize', 'maximize']:
            window.restore()
            speak(f"Undoing {action_type}. Restoring window.")
        
        elif action_type in ['zoom', 'zoom_back', 'move_left', 'move_right']:
            old_box = action[2]
            left, top, width, height = old_box
            window.moveTo(left, top)
            window.resizeTo(width, height)
            speak(f"Undoing {action_type}. Restoring window position and size.")
        
        else:
            speak("Cannot abort that action.")
    else:
        speak("No action to abort.")

def volume_control(action, value=None):
    """Control the volume by steps."""
    if value is None:
        value = 5  # Default step if no value given

    steps = int(value)

    if  "increase" in action:
        for _ in range(steps):
            pyautogui.press("volumeup")
        speak(f"Volume increased.")
    elif "decrease" in action:
        steps=10
        for _ in range(steps):
            pyautogui.press("volumedown")
        speak(f"Volume decreased.")
    elif "mute"  in action:
        pyautogui.press("volumemute")
        speak("Volume muted.")
    elif "unmute" in action :
        pyautogui.press("volumemute")
        speak("Volume unmuted.")
    


# ============================
# NLP-based Features and Self-Improvement
# ============================

def improve_understanding(command):
    """Handle vague commands and improve future understanding."""
    known_commands = [
        "open", "close", "search", "play", "volume", "take screenshot", "tell time", "tell date", "tell joke"
    ]
    
    if any(word in command for word in known_commands):
        speak("I understood that. I'll keep improving my understanding.")
        last_actions_stack.append(lambda: improve_understanding(command))  # ✨
    else:
        speak("Sorry, I didn't quite get that. Can you try again with a clearer command?")
        last_actions_stack.append(lambda: improve_understanding(command))  # ✨

def better_error_handling(command):
    """Learn from mistakes and optimize future responses."""
    if "error" in command or "failed" in command:
        speak("I'll make sure to improve my performance next time.")
        last_actions_stack.append(lambda: better_error_handling(command))  # ✨

def running_todo():
    # Absolute path to the Jarvis folder
    jarvis_folder = r"C:\My Certificates\Jarvis"

    # Script inside Jarvis folder to run
    script_to_run = os.path.join(jarvis_folder, "running.py")

    # Run the Python file as subprocess
    subprocess.run(["python", script_to_run])


# ============================
# Continuous Listening and Responding
# ============================
def run_jarvis():
    """Run Jarvis and execute tasks based on commands."""
    greet_user()  # Greeting added here
    global jarvis_active_until
    while True:
        command = listen()
        current_time=time.time()

        if current_time > jarvis_active_until:
            print("Waiting for wake word 'Jarvis'...")
            continue

        # Now process the command
        if command == "":
            continue

        if "open" in command:
            open_application(command)
        elif "close" in command:
            close_application(command)
        elif "reopen" in command:
            open_closed_tab()
        elif "close window" in command or "minimize" in command:
            close_active_window()
        elif "exit" in command:
            exit_jarvis()
        elif "thank you" in command or "thanks" in command:
            respond_to_gratitude()
        elif "repeat" in command or "again" in command:
            repeat_last_action()
        elif "search" in command:
            search_google(command)
        elif "play" in command:
            play_on_youtube(command)
        elif "time" in command:
            tell_time()
        elif "date" in command:
            tell_date()
        elif "screenshot" in command:
            take_screenshot()
        elif "joke" in command:
            tell_joke()
        elif "volume" in command:
            volume_control(command)
        elif "maximize" in command:
            maximize_window()
        elif "minimise" in command:
            minimize_window()
        elif "zoom back" in command:
            zoom_back_window()
        elif "zoom" in command:
            zoom_window()
        elif "left" in command:
            move_window_left()
        elif "right" in command:
            move_window_right()
        elif "abort" in command or "about the task" in command:
                abort_last_action() 
        elif "again" in command:
            repeat_last_action()
        elif "switch to" in command:
            app_name = command.split("switch to")[-1].strip()  # Extract app name from command
            switch_to_window(app_name)  # Call the function to switch to the app
        elif "last app" in command:
            switch_to_last_window()  # Switch to the last used app
        elif "alt tab" in command:
            alt_tab_switch() 
        elif "schedule" in command:
            listen_for_schedule()
        elif "make schedule" in command or "add task" in command:
            add_task()                
        elif "show schedule" in command or "show tasks" in command:
            show_tasks()
        elif "task completed" in command :
            complete_task()
       
        
        elif "increase volume" in command or "volume up" in command:
            volume_control("increase")
            last_actions_stack.append(("volume", "increase"))

        elif "decrease volume" in command or "volume down" in command:
            volume_control("decrease")
            last_actions_stack.append(("volume", "decrease"))

        elif "mute" in command:
            volume_control("mute")
            last_actions_stack.append(("volume", "mute"))

        elif "unmute" in command:
            volume_control("unmute")
            last_actions_stack.append(("volume", "unmute"))
        
        elif "make a schedule" in command or "todo" in command:
            running_todo()
        elif "heyjarvis" in command :
            speak("Hello Ram")
        elif "write" in command or "type" in command:
            text = command.replace("write", "").replace("type", "").strip()
            if is_notepad_active():
                pyautogui.write(text)
                speak(f"Typing: {text}")
            else:
                speak("Notepad is not currently active. Please open it first.")
            
        else:
            speak("I didn't get it,can i search for it ")
            c2=listen()
            if "YES" in c2:
                response = ask_gemini(command)
                speak(response)
            else:
                speak("OK")
                continue

# Run
if __name__ == "__main__":
   run_jarvis()

