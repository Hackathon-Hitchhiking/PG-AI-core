import base64
import os

from loguru import logger

from pptx_manager.main import PPTXManager
from pptx_manager.models import CreateImageFrameOpts, ImageFrameOpts
from tests.agno_manager import get_head_agent, open_sync_client


test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path)

# text_json = pr.get_all_text_frame_json()
# image_json = pr.get_all_image_json()
slide_size = pr.get_slide_size_px()
slide_count = pr.get_slide_count()


def create_image(prompt: str, slide_id: int, opts: ImageFrameOpts) -> str:
    """
    Генерирует изображение с помощью DALL-E 2 на основе предоставленного запроса и размещает его на указанном слайде.

    Args:
        prompt (str): Текстовый запрос для генерации изображения.
        slide_id (int): Идентификатор слайда, на котором будет размещено сгенерированное изображение.
        opts (ImageFrameOpts): Параметры конфигурации для позиционирования и изменения размера сгенерированного изображения.
            - left (float): Расстояние от левого края слайда.
            - top (float): Расстояние от верхнего края слайда.
            - width (float): Ширина изображения.
            - height (float): Высота изображения.

    Returns:
        str: Сообщение о результате операции с подробным описанием созданной фигуры с изображением.
    """
    logger.debug(f'create_image вызвана с параметрами: prompt={prompt}, slide_id={slide_id}, opts={opts}')

    response = open_sync_client.images.generate(
        model='dall-e-2',
        prompt=prompt,
        n=1,
        size='512x512',
        response_format='b64_json',
    )

    b64_data = response.data[0].b64_json

    image_bytes = base64.b64decode(b64_data)

    return pr.create_image_shape(
        slide_id,
        CreateImageFrameOpts(
            left=opts.left,
            top=opts.top,
            width=opts.width,
            height=opts.height,
            image=image_bytes,
        ),
    )


# Initial description of the presentation
# response = head_agent.run(f'presentation:{pr_json}\n\nОпиши содержимое презентации.')

# Infinite dialogue loop
num = 1
while True:
    user_input = input("Введите запрос для изменения презентации (или 'exit' для выхода): ")
    if user_input.lower() == 'exit':
        print('Диалог завершен.')
        break

    head_agent, text_agent, slide_agent, image_agent, figure_agent = get_head_agent()

    text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape]
    slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]
    image_agent.tools = [create_image]
    figure_agent.tools = [
        pr.update_shape_color,
        pr.update_shape_position,
        pr.update_shape_transparency,
        pr.set_shape_rounding,
        pr.add_figure_shape,
    ]

    text_agent.instructions[-1] = f'структура текстовых элементов: {pr.get_all_text_frame_json()}'
    image_agent.instructions[-1] = f'структура картинок в презентации: {pr.get_all_image_json()}'
    figure_agent.instructions[-1] = (f'структура фигур в презентации в презентации: {pr.get_all_figure_frame_json()}',)
    slide_agent.instructions[-1] = f'кол-во слайдов: {pr.get_slide_count()}'

    slide_size = pr.get_slide_size_px()
    slide_count = pr.get_slide_count()

    text_agent.instructions[-2] = (
        f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
    )
    image_agent.instructions[-2] = (
        f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
    )
    slide_agent.instructions[-2] = (
        f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
    )
    figure_agent.instructions[-2] = (
        f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
    )

    response = head_agent.run(f'{user_input}')

    logger.debug(f'image agent metrics = {image_agent.run_response.metrics if image_agent.run_response else 0}')
    logger.debug(f'text agent metrics = {text_agent.run_response.metrics if text_agent.run_response else 0}')
    logger.debug(f'slide agent metrics = {slide_agent.run_response.metrics if slide_agent.run_response else 0}')
    logger.debug(f'figure agent metrics = {figure_agent.run_response.metrics if figure_agent.run_response else 0}')

    head_agent.memory = None

    print(response.content)
    logger.debug(f'formated_tool_calls = {response.formatted_tool_calls}')
    logger.debug(f'tools = {[response.tools for response in response.member_responses if response.tools is not None]}')
    # Save the updated presentation
    pr.save(f'test_{num}.pptx')
    logger.debug(f'Презентация сохранена как test_{num}.pptx')
    num += 1
