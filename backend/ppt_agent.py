import os
        self,
        ppt_path,
        output_folder,
    ):

        prs = Presentation(ppt_path)

        slides_data = []

        os.makedirs(output_folder, exist_ok=True)

        for index, slide in enumerate(prs.slides):

            slide_text = []

            for shape in slide.shapes:

                if hasattr(shape, "text"):
                    slide_text.append(shape.text)

            combined_text = "\n".join(slide_text)

            image_path = os.path.join(
                output_folder,
                f"slide_{index}.png",
            )

            image = Image.new(
                "RGB",
                (1280, 720),
                color=(20, 20, 20),
            )

            draw = ImageDraw.Draw(image)

            draw.text(
                (50, 100),
                combined_text[:1000],
                fill=(255, 255, 255),
            )

            image.save(image_path)

            slides_data.append(
                {
                    "text": combined_text,
                    "image": image_path,
                }
            )

        return slides_data
