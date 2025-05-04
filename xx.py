import queue
import sounddevice as sd
import vosk
import json
import pyttsx3
import requests
import time

# Initialize TTS engine with Windows-specific driver
engine = pyttsx3.init(driverName='sapi5')  # Use 'sapi5' on Windows
engine.setProperty('rate', 180)  # Speed of speech

# Optional: List and set preferred voice
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # You can try voices[1] for female, if available

# Load Vosk speech recognition model
model = vosk.Model(lang="en-us")
q = queue.Queue()

def callback(indata, frames, time, status):
    if status:
        print("Mic Status:", status)
    q.put(bytes(indata))

# Listen and transcribe voice input
def listen():
    with sd.RawInputStream(samplerate=16000, blocksize=8000, dtype='int16',
                           channels=1, callback=callback):
        print("🎙️ Listening...")
        rec = vosk.KaldiRecognizer(model, 16000)
        while True:
            data = q.get()
            if rec.AcceptWaveform(data):
                result = json.loads(rec.Result())
                text = result.get("text", "")
                if text:
                    print(f"You: {text}")
                    return text

# Send prompt to Phi via Ollama API
def ask_phi(prompt):
    try:
        response = requests.post(
            'http://localhost:11434/api/generate',
            json={
                "model": "phi",
                "prompt": prompt,
                "stream": False
            }
        )
        reply = response.json().get('response', '').strip()
        return reply
    except Exception as e:
        return "Sorry, I couldn't get a response from the AI model."

# Speak out the response
def speak(text):
    print(f"Jarvis: {text}")
    engine.say(text)
    engine.runAndWait()
    time.sleep(0.1)  # Optional: smooth playback in some environments

# Main assistant loop
def main():
    while True:
        try:
            command = listen()
            if "exit" in command.lower() or "quit" in command.lower():
                speak("Goodbye Ram!")
                break
            response = ask_phi(command)
            speak(response)
        except KeyboardInterrupt:
            break
        except Exception as e:
            print("Error:", e)
            speak("Something went wrong.")

# Entry point
if __name__ == "__main__":
    speak("Hello Ram! I am ready to talk.")
    main()
