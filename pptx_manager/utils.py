from pptx.shapes.base import BaseShape
from pptx.slide import Slide


def get_all_methods(cls):
    return [method for method in dir(cls) if not method.startswith('__')]


def hex_to_rgb(hex: str) -> tuple[int, ...]:
    hex_code = hex.lstrip('#')

    return tuple(int(hex_code[i : i + 2], 16) for i in (0, 2, 4))


def get_slide_from_shape(shape: BaseShape) -> Slide:
    for i in range(100):
        shape = shape._parent
        if isinstance(shape, Slide):
            return shape
    raise ValueError('shape is not in a slide')


class CustomList:
    def __init__(self):
        self._data = []

    def insert(self, index, value):
        # If index is larger than the current length, fill the "gap" with None
        if index > len(self._data):
            self._data.extend([None] * (index - len(self._data)))

        # Perform normal list insert (shift everything at/after 'index' to the right)
        self._data.insert(index, value)

    def __getitem__(self, index):
        return self._data[index]

    def __setitem__(self, index, value):
        # If index is beyond length, fill up to that index with None
        if index >= len(self._data):
            self._data.extend([None] * (index - len(self._data) + 1))
        self._data[index] = value

    def __len__(self):
        return len(self._data)

    def __repr__(self):
        return repr(self._data)
