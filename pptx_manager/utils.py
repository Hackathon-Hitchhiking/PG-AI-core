def get_all_methods(cls):
    return [method for method in dir(cls) if not method.startswith('__')]


def hex_to_rgb(hex: str) -> tuple[int, ...]:
    hex_code = hex.lstrip('#')

    return tuple(int(hex_code[i : i + 2], 16) for i in (0, 2, 4))
