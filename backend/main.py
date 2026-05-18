import os
    return {
        "message": "PPT Video Agent API"
    }


@app.post("/generate-video")

async def generate_video(
    file: UploadFile = File(...),
):

    ppt_path = os.path.join(
        UPLOAD_DIR,
        file.filename,
    )

    with open(ppt_path, "wb") as f:
        f.write(await file.read())

    ppt_agent = PPTAgent()

    slides = ppt_agent.extract_slides(
        ppt_path,
        OUTPUT_DIR,
    )

    voice_agent = VoiceAgent(
        ELEVENLABS_API_KEY,
        ELEVENLABS_VOICE_ID,
    )

    for index, slide in enumerate(slides):

        audio_path = os.path.join(
            OUTPUT_DIR,
            f"audio_{index}.mp3",
        )

        voice_agent.generate_audio(
            slide["text"],
            audio_path,
        )

        slide["audio"] = audio_path

    video_agent = VideoAgent()

    final_video = os.path.join(
        OUTPUT_DIR,
        "final_video.mp4",
    )

    video_agent.create_video(
        slides,
        final_video,
    )

    return {
        "video": final_video
    }
