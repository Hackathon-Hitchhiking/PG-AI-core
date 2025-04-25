import base64
import json
import os

from getpass import getpass
from textwrap import dedent

import httpx

from agno.agent import Agent
from agno.media import Image
from agno.models.openai import OpenAIChat
from dotenv import load_dotenv
from loguru import logger
from openai import AsyncOpenAI, OpenAI
from pydantic import BaseModel

from pptx_manager.main import PPTXManager
from pptx_manager.models import CreateImageFrameOpts, ImageFrameOpts, TextFrameOpts
from tests.agno_manager import get_head_agent


load_dotenv()

if not os.environ.get('OPENAI_API_KEY'):
    os.environ['OPENAI_API_KEY'] = getpass('Enter API key for OpenAI: ')


def str2bool(value: str) -> bool:
    return str(value).lower() in ('true', '1', 'yes', 'y', 'on')


if str2bool(os.environ.get('USE_PROXY_URLS', 'True')):
    http_async_client = httpx.AsyncClient(proxy='http://127.0.0.1:1080')
    http_sync_client = httpx.Client(proxy='http://127.0.0.1:1080')
else:
    http_async_client = None
    http_sync_client = None

open_async_client = AsyncOpenAI(
    http_client=http_async_client,
)

open_sync_client = OpenAI(
    http_client=http_sync_client,
)


class CreateStyleAgentImageFrameOpts(ImageFrameOpts):
    prompt: str


class StyleOutput(BaseModel):
    instructions: list[str]


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

    # Append minimalistic style and no text requirements to the prompt
    enhanced_prompt = f'{prompt}. The image should be in a minimalistic style and contain no text.'
    # TODO fix here error
    # ERROR
    # Error
    # code: 400 - {'error': {'message': 'Error in request. Please check
    #                        your input.', 'type': 'invalid_request_error', 'param': None, 'code':
    #                            None}}
    response = open_sync_client.images.generate(
        model='dall-e-3',
        prompt=enhanced_prompt,
        n=1,
        size='1024x1024',
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


STYLE_AGENT_CORE_INSTRUCTIONS = [
    # Analyze the reference presentation for design language, layout, and content structure.
    'Analyze the reference presentation to extract its visual style, layout patterns, and content organization.',
    'Identify color schemes, typography, spacing, and recurring design elements.',
    'Determine how new content should be integrated to match the existing style.',
    # Detect and analyze slide background colors
    'Detect and analyze the background color of each slide in the reference presentation.',
    'Consider how background colors affect readability and visual hierarchy.',
    'Ensure new content maintains appropriate contrast with the background color.',
    'When recommending new slides, specify appropriate background colors that match the presentation style.',
    # Generate actionable instructions for HeadAgent.
    'For the given user request, generate a detailed list of instructions describing:',
    '- Which slides to add or modify.',
    '- Where to place each content element (text, images, figures, etc.) with coordinates and sizes.',
    '- What content to include (text, images, figures, background, etc.).',
    f'- Any specific style requirements ({TextFrameOpts.model_fields.keys()}).',
    '- Background colors for new slides that match the presentation style.',
    'Output only a list of clear, step-by-step instructions for HeadAgent to execute.',
    'Do not perform any slide or image creation yourself.',
    'Ensure instructions are unambiguous and cover all necessary details for implementation.',
    # Enhanced image generation instructions
    'For each slide, actively consider if it would benefit from relevant images that enhance the content:',
    '- Images should complement the text content and reinforce key messages',
    '- Consider using images for abstract concepts, data visualization, or illustrative examples',
    '- Maintain visual consistency with the presentation style',
    '- Ensure images have appropriate contrast with the slide background color',
    'When a slide would benefit from an image, include a "create_image" step for HeadAgent:',
    '- Specify the target slide_id.',
    '- Give a detailed prompt for DALL·E 2 **in English** so the model understands style and subject clearly.',
    '- Include specific visual elements, style, composition, and color palette in the prompt',
    "- Match the prompt's style (palette, mood, level of abstraction) to the reference presentation.",
    "- When generating the prompt, explicitly include instructions to use colors that comply with the presentation's color scheme",
    "- Consider the slide's content, purpose, and background color when crafting the image prompt",
    '- Provide exact left, top, width, height (pixels) for the image frame.',
    'For technical or data-heavy slides, consider images that:',
    '- Visualize complex concepts or processes',
    '- Illustrate technical components or systems',
    '- Represent data trends or statistics in a visual format',
    'For conceptual or strategic slides, consider images that:',
    '- Evoke the right emotional response',
    '- Use metaphors or symbols to represent abstract ideas',
    '- Reinforce the key message or theme of the slide',
    # Figure creation instructions
    'You can also recommend adding geometric shapes and figures to enhance slide design and visual organization:',
    '- Consider using shapes to highlight key information, create visual hierarchy, or organize content',
    '- Shapes can be used for backgrounds, borders, dividers, callouts, or decorative elements',
    '- Ensure shapes match the presentation style and color scheme',
    'When a slide would benefit from shapes or figures, include a "add_figure_shape" step for HeadAgent:',
    '- Specify the target slide_id',
    '- Indicate the shape type (rectangle, rounded rectangle, oval, etc.)',
    '- Provide exact left, top, width, height (pixels) for the shape',
    '- Always specify the position (left, top) and dimensions (height, width) of each figure',
    '- Specify fill color, line color, line width, and transparency, rounding, rotation as needed',
    '- For rounded rectangles, specify the rounding value (0.0-1.0)',
    "- Consider how shapes can be used to create visual structure and guide the viewer's attention",
    # Font and color copying instructions
    'Always copy and use colors and fonts from the template presentation:',
    '- Analyze and extract the exact font names, sizes, and styles used in the template',
    '- Identify and use the color palette from the template (background colors, text colors, accent colors)',
    '- When specifying text properties, always indicate the exact font name and color copied from the template',
    '- For each text element, clearly specify which font and color from the template you are using',
    '- Maintain consistency with the template by using the same font hierarchy (headings, body text, etc.)',
    '- Ensure all new content matches the visual style of the template presentation',
]

style_agent_model = OpenAIChat(id='gpt-4.1-mini', client=open_sync_client, async_client=open_async_client)

test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path, True)

style_agent = Agent(
    name='Style Agent',
    instructions=STYLE_AGENT_CORE_INSTRUCTIONS
    + [
        f'Reference presentation slide size (pixels): {pr.get_slide_size_px()}',
        f'Reference presentation slide count: {pr.get_slide_count()}',
        f'Reference presentation schema: {pr.get_json_schema()}',
    ],
    model=style_agent_model,
    response_model=StyleOutput,
    monitoring=False,
    telemetry=False,
    debug_mode=True,
)


slide_images = []
for slide_id in range(1, pr.get_slide_count() + 1):
    slide_images.append(Image(content=pr.get_slide_image(slide_id), format='png'))

text_for_new_slide = dedent("""
РЕЗУЛЬТАТ ОТ ИСПОЛЬЗОВАНИЯ СИСТЕМЫ
1. Защита файлов от несанкционированного доступа
2. Снижение экономического ущерба
3. Ограниченный доступ к ключам                logger.debug('changing the font name')
""")

big_text_for_new_slide = dedent("""
Факторы, которые влияют на рынок:
Международная напряженность, которая может нарушать мировые цепочки поставок (более 50% рынка сосредоточено
в странах Азии)
Рост спроса на чипы, который определяется не только ростом спроса на высокопроизводительные вычислительные устройства, но и ростом спроса на потребительскую электронику (смартфоны, ПК и др.)
Высокая стоимость развития локального производства (выражается не только в капитальных затратах
на строительство, закупку оборудования и технологии,
но и в качестве подготовки кадров, задействованных
в производстве)
""")

big2_text_for_new_slide = dedent("""
Добавь новый слайд с данным текстом и добавь картинку
Отличительные особенности
возможность интеграции с ЕСИА для авторизации и соблюдения мер ИБ и 152-ФЗ;
гибкая настройка курсов для разных потребностей и уровня подготовки;
возможность интеграции с внешними системами для обогащения данными и для встраивания в различные бизнес-процессы, например интеграция с системами мониторинга;
безопасность – решение состоит в Реестре Отечественного ПО, стек соответствует требованиям ИБ.
""")

user_request = dedent("""
Группы москвичей по отношению к техническим новинкам
Москвичи-энтузиасты
• Любят тестировать новые сервисы и технологии
• Легко разбираются в новых мобильных
приложениях и сервисах
• Склонны искать решение проблемы, если
сталкиваются с трудностями в использовании
сервиса или технологии. Ими движет любопытство
Москвичи-последователи
• Предпочитают использовать новинку после того, как получат
отзывы и рекомендации от лидеров мнений, друзей или
знакомых, чье мнение они считают авторитетным
• Прагматичны: будут пытаться разобраться в приложениях
или сервисах, если действительно в них нуждаются
или им интересно. Ими движет умеренное любопытство
к технологиям и новинкам
Москвичи-консерваторы
• Не любят пользоваться новыми сервисами и технологиями
• Попытаются использовать новое приложение или сервис,
если возникнет такая необходимость или если будут
вынуждены их использовать. Однако при столкновении
с трудностями, скорее откажутся от использования новинок
""")

delete_request = dedent("""
Удали слайд 6 и на первом слайде удали весь текст и оставь только заголовок.
""")

message = style_agent.run(
    f'Проанализируй этот пользовательский запрос и эталонную презентацию. Сгенерируй подробные инструкции для HeadAgent: {big2_text_for_new_slide}',
    # images=slide_images,
)

style_output: StyleOutput = message.content

logger.debug(f'StyleAgent instructions:\n{json.dumps(style_output.model_dump(), indent=4, ensure_ascii=False)}')

head_agent, text_agent, slide_agent, image_agent, figure_agent = get_head_agent()

text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape]  # pr.delete_text_shape
slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]  # pr.delete_slide
image_agent.tools = [create_image, pr.delete_image_shape]
figure_agent.tools = [
    pr.update_shape_color,
    pr.update_shape_position,
    pr.update_shape_transparency,
    pr.set_shape_rounding,
    pr.add_figure_shape,
    pr.delete_figure_shape,
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

head_agent.run(style_output.instructions)

logger.debug(f'metrics = {style_agent.run_response.metrics}')

logger.debug(f'head agent metrics = {head_agent.run_response.metrics if head_agent.run_response else 0}')
logger.debug(f'image agent metrics = {image_agent.run_response.metrics if image_agent.run_response else 0}')
logger.debug(f'text agent metrics = {text_agent.run_response.metrics if text_agent.run_response else 0}')
logger.debug(f'slide agent metrics = {slide_agent.run_response.metrics if slide_agent.run_response else 0}')
logger.debug(f'figure agent metrics = {figure_agent.run_response.metrics if figure_agent.run_response else 0}')

pr.save('style_test.pptx')
