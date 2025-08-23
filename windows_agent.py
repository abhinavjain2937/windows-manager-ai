import google.generativeai as genai
import speech_recognition as sr
import subprocess
import winsound
import time
import win32com.client as wincl
import os

# --- Configuration ---
# Remember to put your secret API key here
GOOGLE_API_KEY = ("AIzaSyCCGRI9rVtGui0j24QCcidUnSRQMlkwcIY")
try:
    genai.configure(api_key=GOOGLE_API_KEY)
except Exception as e:
    print(f"FATAL: Gemini API configuration failed. Is your key valid? Error: {e}")
    exit()

# --- NEW (Plan B): Manual App Paths ---
# If the agent can't find a specific app, add it here manually.
# Use the app's name in lowercase as the key.
# Example: "my game": r"C:\Games\MyAwesomeGame\launcher.exe"
CUSTOM_APP_PATHS = {
    "example app": r"C:\Path\To\Your\App.exe"
}

# --- Native Windows Voice Setup ---
try:
    speaker = wincl.Dispatch("SAPI.SpVoice")
except Exception as e:
    print(f"FATAL: Failed to initialize Windows voice engine: {e}")
    speaker = None


def speak(text):
    """The agent's voice, using the stable, native Windows SAPI."""
    if speaker:
        speaker.Speak(text)
    else:
        print(f"Agent (Voice Error): {text}")


def listen_for_input(prompt):
    """Speaks a prompt, beeps, and then listens."""
    speak(prompt)
    winsound.Beep(440, 250)

    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.pause_threshold = 1.5
        print(f"[Microphone Active...]")
        try:
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=15)
            text_input = recognizer.recognize_google(audio).lower()
            print(f"Heard: '{text_input}'")
            return text_input
        except (sr.WaitTimeoutError, sr.UnknownValueError):
            return None
        except sr.RequestError:
            speak("I'm having trouble connecting to the speech service.")
            return None


# --- UPGRADED: Smart function to find apps everywhere ---
def find_and_open_app(app_name):
    """Searches custom paths, the Desktop, and the Start Menu for an app."""
    speak(f"Searching for the application: {app_name}")

    app_name_lower = app_name.lower()

    # Step 1: Check the manual override list first (fastest)
    if app_name_lower in CUSTOM_APP_PATHS:
        app_path = CUSTOM_APP_PATHS[app_name_lower]
        speak(f"Found {app_name} in custom paths. Opening it now.")
        os.system(f'start "" "{app_path}"')
        return True

    # Step 2: Search the Desktop and Start Menu (automatic)
    search_paths = [
        os.path.join(os.environ['USERPROFILE'], 'Desktop'),
        os.path.join(os.environ.get('PUBLIC', ''), 'Desktop'),  # Use .get for safety
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ['ALLUSERSPROFILE'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')
    ]

    app_path = None
    for path in search_paths:
        if not os.path.isdir(path):
            continue
        for root, dirs, files in os.walk(path):
            for name in files:
                if app_name_lower in name.lower() and name.endswith(('.lnk', '.exe')):
                    app_path = os.path.join(root, name)
                    break
            if app_path: break
        if app_path: break

    if app_path:
        speak(f"Found {app_name}. Opening it now.")
        os.system(f'start "" "{app_path}"')
        return True
    else:
        speak(f"I'm sorry, I couldn't find {app_name} on your system.")
        return False


def get_cmd_command_from_gemini(user_task):
    """Gets a command from Gemini for general tasks."""
    if not user_task: return None
    speak("I couldn't find that as an app, so I'll ask the AI for a general command.")
    prompt = f"You are a Windows command line expert. Convert the user's request into a single, executable command for cmd.exe. Provide ONLY the command. Request: '{user_task}'"
    try:
        # --- FIX: Corrected a typo from GenerModel to GenerativeModel ---
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        response = model.generate_content(prompt)
        return response.text.strip().replace('`', '')
    except Exception as e:
        print(f"Gemini Error: {e}")
        speak("I had a problem connecting to my AI brain.")
        return None


def execute_generic_command(command):
    """Executes a generic command from Gemini after confirmation."""
    if not command: return
    confirmation = listen_for_input(f"The AI suggests the command: {command}. Should I run it?")
    if confirmation and any(word in confirmation for word in ['yes', 'yeah', 'confirm', 'do it', 'okay']):
        speak("Confirmed. Running the command.")
        try:
            result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True, encoding='utf-8')
            output = result.stdout.strip()
            if output:
                speak("The command ran successfully and produced text output, which is in the terminal.")
                print("\n--- Command Output ---\n" + output + "\n----------------------\n")
            else:
                speak("The command completed successfully.")
        except Exception as e:
            speak("The command failed to run.")
            print(f"Execution Error: {e}")
    else:
        speak("Understood. I have cancelled the command.")


# --- Main Application Loop ---
if __name__ == "__main__":
    speak("AI Agent activated.")
    while True:
        command_text = listen_for_input("What is your command?")

        if command_text:
            if any(word in command_text for word in ["exit", "stop", "quit", "goodbye"]):
                speak("Deactivating now. Goodbye!")
                break

            open_keywords = ["open", "launch", "start", "run"]
            is_open_command = any(command_text.startswith(keyword + " ") for keyword in open_keywords)

            if is_open_command:
                app_name_to_find = command_text.split(' ', 1)[1]
                find_and_open_app(app_name_to_find)
            else:
                cmd_to_run = get_cmd_command_from_gemini(command_text)
                if cmd_to_run:
                    execute_generic_command(cmd_to_run)