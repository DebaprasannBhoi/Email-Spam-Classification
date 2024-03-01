import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
from win32com.client import Dispatch
import pythoncom  # Import pythoncom for COM initialization

def speak(text):
    pythoncom.CoInitialize()
    try:
        speak = Dispatch("SAPI.SpVoice")
        speak.Speak(text)
    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        pythoncom.CoUninitialize()

model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

def main():
    st.title("Email Spam Classification Application")
    st.write("Built with Streamlit & Python")
    activities = ["Classification", "About"]
    choices = st.sidebar.selectbox("Select Activities", activities)
    
    if choices == "Classification":
        st.subheader("Classification")
        msg = st.text_input("Enter a text")
        
        if st.button("Process"):
            data = [msg]
            vec = cv.transform(data).toarray()
            result = model.predict(vec)
            
            if result[0] == 0:
                st.success("This is Not A Spam Email")
                speak("This is Not A Spam Email")
            else:
                st.error("This is A Spam Email")
                speak("This is A Spam Email")

if __name__ == "__main__":
    main()


