from elevenlabs.client import ElevenLabs


class VoiceAgent:

    def __init__(
        self,
        api_key,
        voice_id,
    ):

        self.client = ElevenLabs(
            api_key=api_key
        )

        self.voice_id = voice_id

    def generate_audio(
        self,
        text,
        output_file,
    ):

        audio = self.client.text_to_speech.convert(
            voice_id=self.voice_id,
            text=text,
            model_id="eleven_multilingual_v2",
        )

        with open(output_file, "wb") as f:

            for chunk in audio:
                f.write(chunk)

        return output_file
