import math
from copy import deepcopy
from io import BytesIO

from pptx import Presentation
from pptx.enum.dml.color
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE, MSO_THEME_COLOR_INDEX
from pptx.shapes.autoshape import Shape
from pptx.shapes.group import GroupShape
from pptx.shapes.picture import PictureShape
from pptx.util import Pt


def flatten_presentation(pptx_path):
    """
    Загружает презентацию из pptx_path,
    рекурсивно раскрывает все группы (GroupShape) на каждом слайде,
    учитывая масштаб, поворот, flipH/flipV, и т.д.
    Возвращает объект Presentation уже без групп.
    """
    prs = Presentation(pptx_path)

    for slide in prs.slides:
        ungroup_all_in_slide(slide)

    return prs


def ungroup_all_in_slide(slide):
    """
    Ищет все группы на данном слайде и рекурсивно их "сплющивает".
    Повторяет, пока групп не останется.
    """
    while True:
        group_shapes = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.GROUP]
        if not group_shapes:
            break
        ungroup_shape(slide, group_shapes[0])


def ungroup_shape(slide, group_shape):
    """
    "Сплющивает" (раскрывает) один GroupShape:
      - Парсит трансформацию (масштаб, поворот, flipH, flipV, оффсеты)
      - Применяет её к каждой вложенной фигуре
      - Копирует вложенные фигуры на уровень slide.shapes
      - Если находим вложенные группы, рекурсивно раскрываем
      - Удаляем сам group_shape из slide.shapes
    """
    group_xfrm = parse_group_xfrm(group_shape._element)
    sub_shapes = list(group_shape.shapes)

    for sub in sub_shapes:
        sub_xfrm = parse_shape_xfrm(sub._element)
        combined = combine_xfrm(group_xfrm, sub_xfrm)
        copy_shape_with_transform(slide, sub, combined)

    # Удаляем исходную группу из slide.shapes
    slide.shapes._spTree.remove(group_shape._element)


def parse_group_xfrm(grpSp_element):
    """
    Извлекает <p:grpSpPr>/<a:xfrm> для группы,
    возвращает словарь с полями off_x, off_y, scale_x, scale_y, rot_rad, flipH, flipV и т.д.
    """
    nsmap = grpSp_element.nsmap
    grpSpPr = grpSp_element.find("p:grpSpPr", nsmap)
    if grpSpPr is None:
        return identity_xfrm()

    xfrm = grpSpPr.find("a:xfrm", nsmap)
    if xfrm is None:
        return identity_xfrm()

    return parse_xfrm_element(xfrm)


def parse_shape_xfrm(sp_element):
    """
    Извлекает <p:spPr>/<a:xfrm> (или у <p:pic>, <p:graphicFrame>) для обычной фигуры,
    возвращает словарь трансформации.
    """
    nsmap = sp_element.nsmap
    # пытаемся найти spPr (или аналоги)
    spPr = None
    for tag_name in ["p:spPr", "p:picPr", "p:cxnSpPr", "p:graphicFrameLocks", "p:grpSpPr"]:
        tmp = sp_element.find(tag_name, nsmap)
        if tmp is not None:
            spPr = tmp
            break
    if spPr is None:
        xfrm = sp_element.find(".//a:xfrm", nsmap)
        if xfrm is None:
            return identity_xfrm()
        return parse_xfrm_element(xfrm)

    xfrm = spPr.find("a:xfrm", nsmap)
    if xfrm is None:
        return identity_xfrm()

    return parse_xfrm_element(xfrm)


def parse_xfrm_element(xfrm):
    """
    Парсит <a:xfrm>, извлекает off/ext, chOff/chExt, rot, flipH, flipV.
    Возвращает словарь с off_x, off_y, scale_x, scale_y, rot_rad и др.
    """
    def emu(val):
        return int(val) if val else 0

    off = xfrm.find("a:off", xfrm.nsmap)
    ext = xfrm.find("a:ext", xfrm.nsmap)
    chOff = xfrm.find("a:chOff", xfrm.nsmap)
    chExt = xfrm.find("a:chExt", xfrm.nsmap)

    flipH = (xfrm.get("flipH") == "1")
    flipV = (xfrm.get("flipV") == "1")
    rot_val = xfrm.get("rot")
    rot = int(rot_val) if rot_val else 0
    rot_deg = rot / 60000.0
    rot_rad = math.radians(rot_deg)

    off_x = emu(off.get("x")) if off is not None else 0
    off_y = emu(off.get("y")) if off is not None else 0
    ext_cx = emu(ext.get("cx")) if ext is not None else 0
    ext_cy = emu(ext.get("cy")) if ext is not None else 0

    ch_off_x = emu(chOff.get("x")) if chOff is not None else 0
    ch_off_y = emu(chOff.get("y")) if chOff is not None else 0
    ch_ext_cx = emu(chExt.get("cx")) if chExt is not None else 1
    ch_ext_cy = emu(chExt.get("cy")) if chExt is not None else 1

    scale_x = float(ext_cx) / float(ch_ext_cx) if ch_ext_cx != 0 else 1.0
    scale_y = float(ext_cy) / float(ch_ext_cy) if ch_ext_cy != 0 else 1.0

    return {
        "off_x": off_x, "off_y": off_y,
        "ext_cx": ext_cx, "ext_cy": ext_cy,
        "ch_off_x": ch_off_x, "ch_off_y": ch_off_y,
        "ch_ext_cx": ch_ext_cx, "ch_ext_cy": ch_ext_cy,
        "flipH": flipH, "flipV": flipV,
        "rot_rad": rot_rad,
        "scale_x": scale_x, "scale_y": scale_y,
    }


def identity_xfrm():
    """Возвращает "единичную" трансформацию (нет смещения, поворота, флипа, масштаб = 1)."""
    return {
        "off_x": 0, "off_y": 0,
        "ext_cx": 0, "ext_cy": 0,
        "ch_off_x": 0, "ch_off_y": 0,
        "ch_ext_cx": 1, "ch_ext_cy": 1,
        "flipH": False, "flipV": False,
        "rot_rad": 0.0,
        "scale_x": 1.0, "scale_y": 1.0,
    }


def combine_xfrm(parent, child):
    """
    Комбинирует (умножает) трансформации parent (группы) и child (фигуры).
    Итог: масштаб, поворот, flip, смещение.
    """
    scale_x = parent["scale_x"] * child["scale_x"]
    scale_y = parent["scale_y"] * child["scale_y"]
    rot_rad = parent["rot_rad"] + child["rot_rad"]
    flipH = parent["flipH"] ^ child["flipH"]  # XOR
    flipV = parent["flipV"] ^ child["flipV"]

    # Применяем parent-трансформацию к child.off:
    px, py = transform_point(
        child["off_x"], child["off_y"],
        parent["scale_x"], parent["scale_y"],
        parent["rot_rad"],
        parent["flipH"], parent["flipV"]
    )
    off_x = parent["off_x"] + px
    off_y = parent["off_y"] + py

    return {
        "off_x": off_x,
        "off_y": off_y,
        "scale_x": scale_x,
        "scale_y": scale_y,
        "rot_rad": rot_rad,
        "flipH": flipH,
        "flipV": flipV,
        "ext_cx": 0, "ext_cy": 0,
        "ch_off_x": 0, "ch_off_y": 0,
        "ch_ext_cx": 1, "ch_ext_cy": 1,
    }


def transform_point(x, y, scale_x, scale_y, rot_rad, flipH, flipV):
    """
    Применяет flipH, flipV, затем масштаб, потом поворот (rot_rad) к точке (x,y).
    """
    if flipH:
        x = -x
    if flipV:
        y = -y
    x *= scale_x
    y *= scale_y
    cosA = math.cos(rot_rad)
    sinA = math.sin(rot_rad)
    rx = x*cosA - y*sinA
    ry = x*sinA + y*cosA
    return rx, ry


def copy_shape_with_transform(slide, src_shape, xfrm):
    """
    Копирует фигуру src_shape на верхний уровень slide.shapes, применяя трансформацию xfrm.
    """
    base_w = src_shape.width
    base_h = src_shape.height
    new_w = int(round(base_w * xfrm["scale_x"]))
    new_h = int(round(base_h * xfrm["scale_y"]))

    new_left = int(round(xfrm["off_x"]))
    new_top = int(round(xfrm["off_y"]))
    new_rotation_degs = math.degrees(xfrm["rot_rad"])

    stype = src_shape.shape_type

    if stype in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.PLACEHOLDER):
        # AutoShape, Placeholder => add_shape
        auto_type = getattr(src_shape, 'auto_shape_type', None)
        if auto_type is not None:
            new_shape = slide.shapes.add_shape(auto_type, new_left, new_top, new_w, new_h)
            copy_autoshape_formatting(src_shape, new_shape)
            new_shape.rotation = new_rotation_degs
        else:
            # не получилось определить auto_shape_type -> raw xml
            copy_raw_xml_shape(slide, src_shape, new_left, new_top, new_w, new_h, new_rotation_degs)

    elif stype == MSO_SHAPE_TYPE.TEXT_BOX:
        # Для TextBox - та же логика, но add_textbox
        new_shape = slide.shapes.add_textbox(new_left, new_top, new_w, new_h)
        copy_autoshape_formatting(src_shape, new_shape)
        new_shape.rotation = new_rotation_degs

    elif stype == MSO_SHAPE_TYPE.PICTURE:
        if isinstance(src_shape, PictureShape):
            blob = src_shape.image.blob
            new_pic = slide.shapes.add_picture(BytesIO(blob), new_left, new_top, new_w, new_h)
            new_pic.rotation = new_rotation_degs
            # копируем crop
            new_pic.crop_left = src_shape.crop_left
            new_pic.crop_right = src_shape.crop_right
            new_pic.crop_top = src_shape.crop_top
            new_pic.crop_bottom = src_shape.crop_bottom
        else:
            # fallback
            copy_raw_xml_shape(slide, src_shape, new_left, new_top, new_w, new_h, new_rotation_degs)

    elif stype in (
        MSO_SHAPE_TYPE.CHART,
        MSO_SHAPE_TYPE.TABLE,
        MSO_SHAPE_TYPE.SMART_ART,
        MSO_SHAPE_TYPE.LINE,
        MSO_SHAPE_TYPE.CONNECTOR,
        MSO_SHAPE_TYPE.MEDIA,
        MSO_SHAPE_TYPE.CONTENT_PART,
        MSO_SHAPE_TYPE.DIAGRAM,
        MSO_SHAPE_TYPE.WEB_VIDEO,
        MSO_SHAPE_TYPE.INK,
        MSO_SHAPE_TYPE.INK_COMMENT,
        MSO_SHAPE_TYPE.CALLOUT,
        MSO_SHAPE_TYPE.CANVAS,
        MSO_SHAPE_TYPE.FREEFORM,
        MSO_SHAPE_TYPE.COMMENT,
    ):
        # Копируем «сыро» XML
        copy_raw_xml_shape(slide, src_shape, new_left, new_top, new_w, new_h, new_rotation_degs)

    elif stype == MSO_SHAPE_TYPE.GROUP:
        # Вложенная группа. Создаём group_shape и потом ungroup
        nested_grp = slide.shapes.add_group_shape()
        nested_grp.left = new_left
        nested_grp.top = new_top
        nested_grp.width = new_w
        nested_grp.height = new_h
        nested_grp.rotation = new_rotation_degs

        # переносим XML-дочерние
        for nested_sub in list(src_shape.shapes):
            nested_grp.shapes._spTree.append(nested_sub._element)

        ungroup_shape(slide, nested_grp)

    else:
        # fallback
        copy_raw_xml_shape(slide, src_shape, new_left, new_top, new_w, new_h, new_rotation_degs)


def copy_autoshape_formatting(src_shape, dst_shape):
    """
    Базовое копирование заливки, линии, текста (при необходимости расширяйте).
    """
    fill = src_shape.fill
    if fill.type == MSO_FILL_TYPE.SOLID:
        dst_shape.fill.solid()
        if fill.fore_color.type == MSO_THEME_COLOR_INDEX:
            dst_shape.fill.fore_color.theme_color = fill.fore_color.theme_color
            dst_shape.fill.fore_color.brightness = fill.fore_color.brightness
        else:
            dst_shape.fill.fore_color.rgb = fill.fore_color.rgb
    elif fill.type is None:
        dst_shape.fill.background()

    if src_shape.line:
        if src_shape.line.width:
            dst_shape.line.width = src_shape.line.width
        if src_shape.line.color and src_shape.line.color.rgb:
            dst_shape.line.color.rgb = src_shape.line.color.rgb

    if src_shape.has_text_frame:
        dst_shape.text = src_shape.text


def copy_raw_xml_shape(slide, src_shape, new_left, new_top, new_w, new_h, rot_degs=0.0):
    """
    "Сырой" (XML) способ копирования для типов фигур, которые нельзя пересоздать через add_*.
    Меняем координаты, размер, rot в <a:xfrm>, вставляем в slide.shapes._spTree.
    """
    cloned = deepcopy(src_shape._element)
    nsmap = cloned.nsmap
    xfrm = cloned.find(".//a:xfrm", nsmap)
    if xfrm is not None:
        off = xfrm.find("a:off", nsmap)
        ext = xfrm.find("a:ext", nsmap)
        if off is not None:
            off.set("x", str(new_left))
            off.set("y", str(new_top))
        if ext is not None:
            ext.set("cx", str(new_w))
            ext.set("cy", str(new_h))

        # rot_degs -> rot_attr в 1/60000 градуса
        rot_attr = int(round(rot_degs * 60000))
        xfrm.set("rot", str(rot_attr))

    slide.shapes._spTree.append(cloned)



# ------------------------
# ПРИМЕР ИСПОЛЬЗОВАНИЯ
# ------------------------
pptx_file = "../test_data/test_dit.pptx"
prs = flatten_presentation(pptx_file)
prs.save("flattened_no_groups.pptx")