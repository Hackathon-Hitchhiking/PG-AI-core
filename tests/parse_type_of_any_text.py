from pptx import Presentation


def find_text_recursive(shapes, target_text):
    for shape in shapes:
        # Check text in shape
        if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
            # Check full shape text
            if shape.text and shape.text.strip() == target_text:
                print(f"Shape name: '{getattr(shape, 'name', None)}', type: {type(shape)} (shape.text)")
                return True
            # Check paragraphs/runs for exact match (sometimes text is split)
            for paragraph in shape.text_frame.paragraphs:
                # Check paragraph text
                if paragraph.text and paragraph.text.strip() == target_text:
                    print(f"Shape name: '{getattr(shape, 'name', None)}', type: {type(shape)} (paragraph.text)")
                    return True
                for run in paragraph.runs:
                    if run.text and run.text.strip() == target_text:
                        print(f"Shape name: '{getattr(shape, 'name', None)}', type: {type(shape)} (run.text)")
                        return True
        # Check for grouped shapes or tables
        if hasattr(shape, 'shapes'):
            if find_text_recursive(shape.shapes, target_text):
                return True
        # Check for table cells
        if hasattr(shape, 'table'):
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    if cell.text and cell.text.strip() == target_text:
                        print(f"Table cell in shape '{getattr(shape, 'name', None)}', type: {type(shape)} (cell.text)")
                        return True
    return False


pptx_path = 'test_data/final_test_1.pptx'
prs = Presentation(pptx_path)
slide = prs.slides[1]

found = find_text_recursive(slide.shapes, '3')
if not found:
    # Try also in placeholders (slide numbers are often placeholders)
    for placeholder in slide.placeholders:
        if hasattr(placeholder, 'text') and placeholder.text and placeholder.text.strip() == '1':
            print(
                f"Placeholder name: '{getattr(placeholder, 'name', None)}', type: {type(placeholder)} (placeholder.text)"
            )
            found = True
            break
if not found:
    print('Not found')
