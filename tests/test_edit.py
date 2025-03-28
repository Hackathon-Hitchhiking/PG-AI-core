from parsers.pptx_editor import PPTXEditor


if __name__ == '__main__':
    pptx_editor = PPTXEditor()

    # Изменение заднего фона на зелёный
    pptx_editor.change_background_color(
        pptx_path='../test_data/test_dit.pptx',
        output_pptx_path='test_sources/green_background.pptx',
        color_hex='#00FF00',
    )

    # Изменение цвета графиков на красный
    pptx_editor.change_chart_colors(
        pptx_path='test_sources/green_background.pptx',
        output_pptx_path='test_sources/red_charts.pptx',
        color_hex='#FF0000',
    )
