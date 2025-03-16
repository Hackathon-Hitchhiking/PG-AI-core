import json
from parsers.pptx import PPTXHandler

if __name__ == "__main__":
    handler = PPTXHandler(template_json_path="test_sources/template.json")

    presentation_info = handler.parse_presentation(
        pptx_path="test_sources/test_dit.pptx",
        output_image_dir="test_sources/extracted_images"
    )

    handler.create_presentation_from_template(
        presentation_info=presentation_info,
        output_pptx_path="test_sources/output_presentation.pptx"
    )

    print("✅ Новая презентация успешно создана!")
