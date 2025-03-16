from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor
def rgb_to_hex(rgb_color: RGBColor):
    return '#{:02x}{:02x}{:02x}'.format(rgb_color[0], rgb_color[1], rgb_color[2])


def hex_to_rgb(hex_color: str):
    hex_color = hex_color.lstrip('#')
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )

def resolve_chart_type(chart_type_str):
    chart_type_mapping = {
        "COLUMN_CLUSTERED": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "COLUMN_STACKED": XL_CHART_TYPE.COLUMN_STACKED,
        "COLUMN_STACKED_100": XL_CHART_TYPE.COLUMN_STACKED_100,
        "BAR_CLUSTERED": XL_CHART_TYPE.BAR_CLUSTERED,
        "BAR_STACKED": XL_CHART_TYPE.BAR_STACKED,
        "BAR_STACKED_100": XL_CHART_TYPE.BAR_STACKED_100,
        "LINE": XL_CHART_TYPE.LINE,
        "PIE": XL_CHART_TYPE.PIE,
        "SCATTER": XL_CHART_TYPE.XY_SCATTER,
    }
    return chart_type_mapping.get(chart_type_str.split(" ")[0], XL_CHART_TYPE.COLUMN_CLUSTERED)