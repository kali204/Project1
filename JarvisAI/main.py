import webbrowser
import speech_recognition as sr
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    if text:
        speaker.Speak(f"say {text}")

def listen_for_wake_word():
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()

    # Adjust the energy threshold as needed based on your environment
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)

    print("Listening for the wake word...")

    while True:
        with microphone as source:
            audio = recognizer.listen(source)

        try:
            # Use Google's Speech Recognition API to recognize speech
            # Adjust language and timeout parameters as needed
            text = recognizer.recognize_google(audio, language="en-US")
            print(f"You said: {text}")

            # Check if the wake word is spoken
            if "hey jarvis" in text.lower():
                print("Wake word detected. Your assistant is now listening.")
                return True  # Return True when the wake word is detected

        except sr.UnknownValueError:
            print("Could not understand the audio.")
        except sr.RequestError as e:
            print(f"Error: {str(e)}")

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        print("Listening...")

        try:
            audio = r.listen(source)
            print("Processing audio...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            print("Sorry, I could not understand what you said.")
            return "Sorry, I could not understand."
        except sr.RequestError as e:
            print(f"Request error: {str(e)}")
            return "Sorry, an error occurred during speech recognition."
        except Exception as ex:
            print(f"Error occurred: {str(ex)}")
            return "An unexpected error occurred."

if __name__ == '__main__':
    # Listen for the wake word
    while not listen_for_wake_word():
        pass  # Keep listening until the wake word is detected

    text = takecommand()
    sites = [["YouTube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
             ["Google", "https://www.Google.com"]]
    for site in sites:
        if f"Open {site[0]}".lower() in text.lower():
            say(f"Opening {site[0]} Sir ..")
            webbrowser.open(site[1])
    # say(text)
