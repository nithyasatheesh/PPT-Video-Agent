import os

from moviepy.editor import (
    AudioFileClip,
    ImageClip,
    concatenate_videoclips,
)


class VideoAgent:

    def create_video(
        self,
        slides,
        output_path,
    ):

        clips = []

        for slide in slides:

            image_path = slide["image"]
            audio_path = slide["audio"]

            audio = AudioFileClip(audio_path)

            clip = ImageClip(image_path)

            clip = clip.set_duration(audio.duration)

            clip = clip.set_audio(audio)

            clips.append(clip)

        final_video = concatenate_videoclips(
            clips,
            method="compose",
        )

        final_video.write_videofile(
            output_path,
            fps=24,
            codec="libx264",
            audio_codec="aac",
        )

        return output_path
