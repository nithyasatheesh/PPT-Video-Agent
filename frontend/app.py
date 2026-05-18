import requests
import streamlit as st


st.set_page_config(
    page_title="PPT to Video Generator",
    layout="wide",
)

st.title("🎥 PPT to Video AI Agent")

uploaded_file = st.file_uploader(
    "Upload PPT",
    type=["pptx"],
)

if st.button("Generate Video"):

    if not uploaded_file:
        st.warning("Upload PPT")
        st.stop()

    files = {
        "file": uploaded_file.getvalue()
    }

    response = requests.post(
        "http://127.0.0.1:8000/generate-video",
        files={
            "file": (
                uploaded_file.name,
                uploaded_file.getvalue(),
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
        },
    )

    data = response.json()

    st.success("Video Generated")

    st.video(data["video"])
