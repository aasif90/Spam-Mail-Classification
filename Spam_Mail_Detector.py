import pickle
import streamlit as st
from win32com.client import Dispatch
import pythoncom
from datetime import datetime  # Importing datetime module
import time  # Import time module for sleep functionality

# Initialize pyttsx3 for text-to-speech
def speak(text):
    pythoncom.CoInitialize() 
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(text)

# Load model and vectorizer
model = pickle.load(open("spam.pkl", "rb"))
cv = pickle.load(open("vectorizer.pkl", "rb"))

# Function to initialize history in session_state if not already initialized
def init_history():
    if 'history' not in st.session_state:
        st.session_state.history = []

# Function to display history in sidebar with date and time
def display_history():
    if 'history' in st.session_state and st.session_state.history:
        st.sidebar.subheader("History")
        for idx, item in enumerate(st.session_state.history):
            # Format the datetime object to show the date and time
            timestamp = item['timestamp'].strftime('%Y-%m-%d %H:%M:%S')  # Format as 'YYYY-MM-DD HH:MM:SS'
            st.sidebar.write(f"{idx + 1}. {item['message']} - {item['prediction']} - {timestamp}")

# Prevent repeated audio welcome by using a flag
def play_welcome_audio():
    if 'welcome_played' not in st.session_state:
        st.session_state.welcome_played = True
        speak("Welcome to the Email Spam Classification System")

# Function to simulate a loading screen
def loading_screen():
    with st.spinner("Classifying the email... Please wait!"):
        time.sleep(3)  # Simulate processing time (3 seconds)

def main():
    # Initialize history and load initial welcome message
    init_history()
    play_welcome_audio()  # Play the welcome audio once

    # Display the title and description
    st.title("Email Spam Classification")
    st.subheader("Build By MD AASIF RAZA")

    # Input box for the user to enter a text
    msg = st.text_input("Enter a Text:", key="input_message")  # Use a unique key

    # Display prediction button
    if st.button("Predict"):
        # Show the loading screen
        loading_screen()

        # Perform prediction after the simulated loading
        data = [msg]
        vect = cv.transform(data).toarray()
        prediction = model.predict(vect)
        result = prediction[0]
        
        # Get current date and time
        current_time = datetime.now()  # Capture the current date and time
        
        # Check prediction and update the session state
        if result == 1:
            st.error("This is a spam mail")
            speak("This is a spam mail")
            prediction_result = "Spam"
        else:
            st.success("This is a ham mail")
            speak("This is a Ham mail")
            prediction_result = "Ham"

        # Save the prediction result, message, and timestamp to history
        st.session_state.history.append({
            'message': msg, 
            'prediction': prediction_result, 
            'timestamp': current_time  # Save the current time in the history
        })
        
        # Display the updated history
        display_history()

if __name__ == "__main__":
    main()