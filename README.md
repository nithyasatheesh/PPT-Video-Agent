# PPT-Video-Agent

A single-app Streamlit Community Cloud project that turns uploaded PowerPoint files into short narrated MP4 videos.

## Deployment target

This repository is intentionally optimized for **Streamlit Community Cloud only**:

- No FastAPI backend
- No Render deployment
- No localhost service calls
- No MoviePy or ImageMagick dependency
- Python 3.11 runtime
- Lightweight FFmpeg usage through `imageio-ffmpeg`

## How it works

1. Upload a `.pptx` file in Streamlit.
2. The app extracts slide text with `python-pptx`.
3. The app creates readable 1280x720 slide images with Pillow.
4. OpenAI GPT-4o generates concise narration for each slide.
5. OpenAI TTS generates MP3 narration.
6. A lightweight FFmpeg pipeline combines still slide images and narration into an MP4.
7. Streamlit displays a video preview and download button.

## Streamlit Cloud setup

1. Set the main file path to `app.py`.
2. Add this secret in Streamlit Community Cloud:

```toml
OPENAI_API_KEY = "your_openai_api_key"
```

3. Deploy with the included `runtime.txt` and `requirements.txt`.

## Local run

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Notes

The app renders a clean text-based version of slides rather than attempting pixel-perfect PowerPoint rendering. This avoids LibreOffice, ImageMagick, MoviePy, and platform-specific rendering failures on Streamlit Community Cloud.
