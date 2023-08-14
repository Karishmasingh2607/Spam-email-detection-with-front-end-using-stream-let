import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
from PIL import Image
from win32com.client import Dispatch

def speak(text):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(text)

model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

# Set Streamlit theme
st.set_page_config(
    page_title="Email Spam Classification Application",
    page_icon="📧",
    layout="wide",
)

def main():
    st.title("Email Spam Classification Application")

    # Set background color and style
    st.markdown(
        """
        <style>
        body {
            background-color: #F5F5F5;
            color: #333333;
            font-family: Arial, sans-serif;
        }
        .stButton button {
            background-color: #FF5722;
            color: white;
            border-radius: 5px;
            cursor: pointer;
            font-size: 18px;
            padding: 10px 20px;
            border: none;
        }
        .stTextInput input {
            width: 80%;
            padding: 10px;
            font-size: 18px;
            margin: auto;
            margin-top: 20px;
            border: 1px solid #ccc;
            border-radius: 5px;
        }
        .stWarning {
            font-size: 18px;
            color: #FF5722;
        }
        .stSuccess {
            font-size: 18px;
            color: #4CAF50;
        }
        .stError {
            font-size: 18px;
            color: #FF5722;
        }
        .title {
            font-size: 36px;
            text-align: center;
            color: #FF5722;
            margin-top: 20px;
        }
        .subheader {
            font-size: 24px;
            text-align: center;
            margin-top: 20px;
        }
        .result {
            font-size: 28px;
            text-align: center;
            margin-top: 20px;
            animation: pulse 1s infinite;
        }
        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.05); }
            100% { transform: scale(1); }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown('<p class="title">Email Spam Classification Application</p>', unsafe_allow_html=True)

    activites = ["Classification", "About"]
    choices = st.sidebar.selectbox("Select Activities", activites)

    if choices == "Classification":
        st.subheader("Classification")
        msg = st.text_input("Enter the content of the mail", value="")

        if st.button("Process"):
            if msg:
                data = [msg]
                vec = cv.transform(data).toarray()
                result = model.predict(vec)
                if result[0] == 0:
                    st.success("This is Not A Spam Email")
                    speak("This is Not A Spam Email")
                else:
                    st.error("This is A Spam Email")
                    speak("This is A Spam Email")
            else:
                st.warning("Please enter the content of the mail.")

    st.image("spam.gif", use_column_width=True)

    with st.spinner("Performing classification..."):
        st.success("Classification completed!")
        result_text = st.empty()
        result_text.markdown('<p class="result">Result will appear here</p>', unsafe_allow_html=True)
       
        if 'result' in locals():
            if result[0] == 0:
                result_text.success("This is Not A Spam Email")
            else:
                result_text.error("This is A Spam Email")

if __name__ == "__main__":
    main()
