import copy

from collections import defaultdict

from pptx_manager.main import PPTXManager


def process_presentation(data):
    def create_signature(element, element_type):
        signature = {
            'left': round(element['left'], 2),
            'top': round(element['top'], 2),
            'width': round(element['width'], 2),
            'height': round(element['height'], 2),
        }
        if element_type == 'text':
            style_keys = [
                'font_name',
                'font_size',
                'bold',
                'italic',
                'underline',
                'strike',
                'color',
                'align',
                'vertical_align',
            ]
            for key in style_keys:
                value = element.get(key)
                if hasattr(value, 'value'):
                    value = value.value
                if isinstance(value, list):
                    value = tuple(value)
                signature[key] = value
        elif element_type == 'figure':
            style_keys = ['shape_type', 'color', 'line_color', 'line_width', 'rounding', 'rotation']
            for key in style_keys:
                value = element.get(key)
                if isinstance(value, list):
                    value = tuple(value)
                signature[key] = value
        elif element_type == 'image':
            pass
        return tuple(sorted(signature.items()))

    element_counter = defaultdict(set)
    slide_keys = [k for k in data.keys() if isinstance(k, int)]
    for slide_num in slide_keys:
        slide = data[slide_num]
        for element_type in ['text', 'image', 'figure']:
            for element in slide.get(element_type, []):
                sig = (element_type, create_signature(element, element_type))
                element_counter[sig].add(slide_num)

    total_slides = len(slide_keys)
    repeated_signatures = {sig for sig, slides in element_counter.items() if len(slides) == total_slides}

    first_slide = data[slide_keys[0]]
    result = {'text': [], 'image': [], 'figure': [], 'slide': copy.deepcopy(first_slide.get('slide', {}))}
    for element_type in ['text', 'image', 'figure']:
        for element in first_slide.get(element_type, []):
            sig = (element_type, create_signature(element, element_type))
            if sig in repeated_signatures:
                cloned = copy.deepcopy(element)
                if element_type == 'text':
                    cloned['text'] = 'TEXT'
                result[element_type].append(cloned)
    return result


# Пример использования
test_pres_path = 'test_data\\final_test_1.pptx'
pr = PPTXManager(test_pres_path, False)

data = pr.get_json_schema()
# print(data)
# print(data)
filtered_data = process_presentation(data)
print(filtered_data)
