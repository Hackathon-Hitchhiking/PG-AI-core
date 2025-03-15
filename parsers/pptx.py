import os
import json
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
import textwrap

class PPTXHandler:
    def __init__(self, template_json_path):
        with open(template_json_path, "r", encoding="utf-8") as f:
            self.template = json.load(f)["slide_template"]
        self.content = {}
        self.slide_width = None
        self.slide_height = None

    def hex_to_rgb(self, hex_color):
        """Convert HEX color to RGB."""
        hex_color = hex_color.lstrip('#')
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

    def parse_presentation(self, pptx_path, output_image_dir):
        """Parse an existing presentation and extract content and images."""
        prs = Presentation(pptx_path)
        self.slide_width = prs.slide_width.pt
        self.slide_height = prs.slide_height.pt

        if not os.path.exists(output_image_dir):
            os.makedirs(output_image_dir)

        presentation_info = {
            "dimensions": {
                "width": self.slide_width,
                "height": self.slide_height
            },
            "slides": []
        }

        for slide_index, slide in enumerate(prs.slides):
            slide_info = {"slide_number": slide_index + 1, "shapes": []}

            for shape_index, shape in enumerate(slide.shapes):
                if shape.has_text_frame:
                    shape_info = {
                        "type": "text",
                        "text": shape.text,
                        "coordinates": {
                            "x": shape.left.pt,
                            "y": shape.top.pt,
                            "width": shape.width.pt,
                            "height": shape.height.pt
                        }
                    }
                    slide_info["shapes"].append(shape_info)
                elif shape.shape_type == 13:  # Shape type for images
                    image_filename = f"slide_{slide_index + 1}_image_{shape_index + 1}.png"
                    image_path = os.path.join(output_image_dir, image_filename)
                    with open(image_path, "wb") as img_file:
                        img_file.write(shape.image.blob)

                    shape_info = {
                        "type": "image",
                        "image_path": image_path,
                        "coordinates": {
                            "x": shape.left.pt,
                            "y": shape.top.pt,
                            "width": shape.width.pt,
                            "height": shape.height.pt
                        }
                    }
                    slide_info["shapes"].append(shape_info)

            presentation_info["slides"].append(slide_info)

        self.content = presentation_info

    def create_presentation_from_template(self, output_pptx_path):
        """Create a new presentation based on the parsed content and template."""
        prs = Presentation()
        prs.slide_width = Pt(self.content["dimensions"]["width"])
        prs.slide_height = Pt(self.content["dimensions"]["height"])

        for slide_content in self.content["slides"]:
            slide_layout = prs.slide_layouts[0]  # Blank layout
            slide = prs.slides.add_slide(slide_layout)

            # Apply background color from template
            bg_color_hex = self.template.get("background_color", "#FFFFFF")
            bg_fill = slide.background.fill
            bg_fill.solid()
            bg_fill.fore_color.rgb = self.hex_to_rgb(bg_color_hex)

            for shape in slide_content["shapes"]:
                coords = shape["coordinates"]
                valid_x = max(0, min(coords["x"], prs.slide_width.pt - coords["width"]))
                valid_y = max(0, min(coords["y"], prs.slide_height.pt - coords["height"]))
                valid_width = min(coords["width"], prs.slide_width.pt - valid_x)
                valid_height = min(coords["height"], prs.slide_height.pt - valid_y)

                if shape["type"] == "text":
                    textbox = slide.shapes.add_textbox(
                        Pt(valid_x), Pt(valid_y), Pt(valid_width), Pt(valid_height)
                    )
                    text_frame = textbox.text_frame
                    text_frame.clear()
                    text_frame.word_wrap = True
                    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

                    p = text_frame.add_paragraph()
                    p.text = textwrap.fill(shape.get("text", ""), width=50)
                    p.font.size = Pt(self.template["default_font"]["size"])
                    p.font.name = self.template["default_font"]["name"]
                    p.font.bold = self.template["default_font"]["bold"]
                    p.font.italic = self.template["default_font"]["italic"]
                    p.font.color.rgb = self.hex_to_rgb(self.template["default_font"]["color"])
                    p.alignment = PP_ALIGN.LEFT

                elif shape["type"] == "image":
                    img_path = shape.get("image_path", "")
                    if img_path:
                        slide.shapes.add_picture(img_path, Pt(valid_x), Pt(valid_y), Pt(valid_width), Pt(valid_height))

        prs.save(output_pptx_path)