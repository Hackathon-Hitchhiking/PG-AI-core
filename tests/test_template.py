import base64
import json
import os

from getpass import getpass

import httpx

from agno.agent import Agent
from agno.media import Image
from agno.models.openai import OpenAIChat
from dotenv import load_dotenv
from loguru import logger
from openai import AsyncOpenAI, OpenAI
from pydantic import BaseModel


# Save all logs to test.log
logger.add('test.log', rotation='100 MB', encoding='utf-8')

from pptx_manager.main import PPTXManager
from pptx_manager.models import CreateImageFrameOpts, ImageFrameOpts
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


# --- Подготовка только шаблонной презентации ---
template_pres_path = 'test_data/test_template.pptx'
template_pr = PPTXManager(template_pres_path, True)
template_slide_images = [Image(content=template_pr.get_slide_image(1), format='png')]

# Получаем количество слайдов в основной презентации
test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path, True)
main_slide_count = pr.get_slide_count()
# --- Инструкции для StyleAgent: только шаблон и число слайдов основной ---
STYLE_AGENT_CORE_INSTRUCTIONS = [
    'You are StyleAgent. You are provided ONLY with a template presentation (template.pptx).',
    f'The main presentation is NOT provided. You only know that it currently has {main_slide_count} slides.',
    'Your task: generate a step-by-step, fully technical instruction set for HeadAgent to CREATE a NEW slide as the LAST slide (slide number {main_slide_count + 1}) in the main presentation, using the template slide as a pixel-perfect technical reference.',
    'For each element (shape, text box, title, subtitle, background, etc.) on the template slide, extract and explicitly specify ALL of the following parameters:',
    '- Type of element (e.g., rectangle, rounded rectangle, oval, text box, title, subtitle, background, image, etc.)',
    '- Exact coordinates: left, top, width, height (in pixels)',
    '- For text: font family, font size, font color (RGB or hex), bold/italic/underline, alignment, line spacing, text content, and z-order (layer order)',
    '- For shapes: shape type, fill color, border color, border width, transparency, rounding, rotation, and z-order',
    '- For images: position, size, z-order, and if needed, a prompt for image generation matching the template style',
    '- For background: color (RGB or hex), gradient or pattern if present',
    '- Z-order (layer order) for all elements to ensure correct stacking',
    '- Any additional style or formatting details present in the template',
    'Completely ignore any logos or branding images: do not generate, copy, or describe them in the instructions.',
    'HeadAgent does NOT have access to the template presentation or its images. You must provide all technical details explicitly.',
    "Do NOT use general phrases like 'match the template' or 'use similar style' — always specify exact values and parameters.",
    'If the template slide contains multiple elements, generate a separate instruction for each, with all required parameters.',
    "If the template contains images, and the user wants new content, generate a DALL-E prompt that matches the template's style and color palette.",
    "Replace the template's text content with content relevant to the user's topic, but preserve all style and formatting parameters.",
    'If the template contains placeholder images or icons, generate new ones matching the style, but relevant to the new topic.',
    'Output only a list of clear, step-by-step, technical instructions for HeadAgent to execute.',
    'Each instruction must be self-contained and unambiguous.',
    'Always specify the slide number for each instruction. For this task, the slide number is {main_slide_count + 1}.',
    "Example: 'On slide 8, add a rectangle shape at left=100, top=200, width=400, height=100, fill_color=#F0F0F0, border_color=#000000, border_width=2, transparency=0.1, z_order=1.'",
    'Example: \'On slide 8, add a text box at left=120, top=220, width=360, height=60, font_family=Arial, font_size=28, font_color=#222222, bold=True, italic=False, alignment=center, text="[user content]", z_order=2.\'',
    f'Template presentation slide size (pixels): {template_pr.get_slide_size_px()}',
    f'Template presentation schema: {template_pr.get_json_schema()}',
    'When generating instructions, specify all coordinates, sizes, colors, fonts, and other style details as extracted from the template.',
    'Ensure the new slide visually matches the template as closely as possible, but with new content as requested by the user.',
    'If any parameter is missing in the template, make a reasonable technical assumption and state it explicitly.',
    'Do not perform any slide or image creation yourself.',
]

style_agent_model = OpenAIChat(id='gpt-4.1-mini', client=open_sync_client, async_client=open_async_client)

style_agent = Agent(
    name='Style Agent',
    instructions=STYLE_AGENT_CORE_INSTRUCTIONS,
    model=style_agent_model,
    response_model=StyleOutput,
    monitoring=False,
    telemetry=False,
    debug_mode=True,
)

# User prompt for template-based slide (полная текстовка для слайда)
template_request = (
    'Создай последний слайд в основной презентации, используя шаблонный слайд как технический образец. '
    "Тема слайда: 'Большое будущее нейросетей'. "
    "Заголовок: 'Большое будущее нейросетей'. "
    "Основной текст: 'Нейросети становятся ключевым инструментом для развития технологий, экономики и общества. "
    'Они открывают новые горизонты в медицине, образовании, бизнесе и науке. '
    'Инвестиции в развитие нейросетей обеспечивают конкурентоспособность и инновационное лидерство. '
    "Будущее принадлежит тем, кто использует потенциал искусственного интеллекта для решения глобальных задач.'"
    'Сначала нужно создать последний слайд, а затем добавить в него элементы и информацию, разнесенную по блокам.'
)

# Run StyleAgent только с шаблоном и запросом
message = style_agent.run(
    f'Сгенерируй подробные технические инструкции для HeadAgent по созданию последнего слайда (номер {main_slide_count + 1}) в основной презентации по шаблону, с новой тематикой: {template_request}',
    images=template_slide_images,
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
