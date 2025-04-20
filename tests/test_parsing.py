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
from pydantic import BaseModel, Field

from pptx_manager.main import PPTXManager
from pptx_manager.models import CreateImageFrameOpts, CreateTextFrameOpts, ImageFrameOpts


load_dotenv()

if not os.environ.get('OPENAI_API_KEY'):
    os.environ['OPENAI_API_KEY'] = getpass.getpass('Enter API key for OpenAI: ')


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


class StyleAgentResponse(BaseModel):
    text_blocks: list[CreateTextFrameOpts]
    image_blocks: list[CreateStyleAgentImageFrameOpts]
    slide_index_to_add: int = Field(description='the slide on what setting the information')
    background_color_rgb: list[int]


model = OpenAIChat(id='gpt-4o', client=open_sync_client, async_client=open_async_client)

test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path, True)

style_agent = Agent(
    name='Style Agent',
    instructions=[
        # Core Design Matching Instructions
        "Create new slides that match the existing presentation's design language.",
        'Analyze current slides to replicate visual and structural elements.',
        'Ensure cohesive visual identity across all presentation materials.',
        'Input: textual content for the new slide.',
        'Input: access to all existing slides in the presentation.',
        # Design Analysis Instructions
        'Analyze color schemes: background, text, accents, highlights.',
        'Analyze typography: font families, sizes, weights, styles.',
        'Analyze layout patterns: margins, alignment, spacing.',
        'Analyze visual elements: shapes, lines, icons, decorative elements.',
        'Analyze image styling: borders, shadows, placement conventions.',
        'Identify recurring design patterns across slides.',
        'Identify header and footer designs.',
        'Identify title formatting conventions.',
        'Identify bullet point and list styling.',
        'Identify transitions between content types.',
        'Identify positioning of similar content elements.',
        'Document RGB values for all design elements.',
        # Content Analysis Instructions
        'Analyze provided text for logical structure and hierarchy.',
        'Determine key points for emphasis.',
        'Identify content suitable for visualization.',
        'Evaluate text volume against typical slide density.',
        'Determine slide purpose in the context of the presentation.',
        # Layout and Positioning Instructions
        'Select layout template matching content type and purpose.',
        'Position text based on patterns from reference slides.',
        'Maintain consistent margins and padding.',
        'Follow alignment principles: left, right, center.',
        'Preserve spacing between elements.',
        'Replicate text block dimensions where appropriate.',
        'Specify positioning in pixel coordinates.',
        'Check for and prevent element overlap in all slide layouts.',
        'Implement collision detection between all content blocks.',
        'Maintain minimum spacing between adjacent elements.',
        'Adjust element positioning to eliminate overlapping content.',
        # Typography Instructions
        'Match typography: titles, subtitles, body text.',
        'Replicate font sizes for each hierarchy level.',
        'Apply consistent font weights and styles.',
        'Maintain line height and character spacing.',
        # Color Application Instructions
        'Use identical RGB values for text, backgrounds, accents.',
        'Maintain consistent color relationships.',
        'Apply color coding for emphasis or categorization.',
        # Visual Element Instructions
        'Reproduce standard shapes, lines, separators.',
        'Apply identical effects: shadows, gradients, transparency.',
        'Maintain styling for bullets, numbering, annotations.',
        # Image Handling Instructions
        'Specify image placement coordinates and dimensions.',
        'Provide image generation prompts matching visual style.',
        'Include framing, borders, and effects for images.',
        'Ensure integration of images with surrounding elements.',
        'Create new images when required based on slide content.',
        'Generate appropriate imagery that matches presentation theme.',
        'Select image styles consistent with existing visual language.',
        'Optimize generated images for presentation format and resolution.',
        'Balance image prominence with textual content.',
        # Data Visualization Instructions
        'Match style of existing charts and diagrams.',
        "Use presentation's color palette for visualizations.",
        'Apply consistent labeling and annotation for charts.',
        # Quality Assurance Instructions
        'Verify style consistency with surrounding slides.',
        'Check text hierarchy and readability.',
        'Check visual composition balance.',
        'Verify color accuracy.',
        'Check information density.',
        'Verify spacing and alignment of elements.',
        'Ensure cohesion with presentation flow.',
        'Confirm no elements overlap or obscure other content.',
        'Test readability of all text elements.',
        'Validate that all generated images display properly.',
        # Technical Parameters
        'Use pixel units for all measurements.',
        'Use RGB format for color specifications.',
        'Use consistent coordinate references.',
        'Match font sizes with existing slides.',
        'Consider slide number for transitions.',
        'Apply standard margin buffer between all content elements.',
        'Follow z-index hierarchy for layered elements.'
        # f'параметры, которые можно использовать для текстового блока {TextFrameOpts.model_fields.keys()}',
        # f'параметры, которые нужно использовать для размещение всех объектов на слайде {ImageFrameOpts.model_fields.keys()}, используй их вместе со всеми блоками',
        f'take into account that the size of the slide is equal {pr.get_slide_size_px()} in pixels when you place objects',
        f'take into account that the number of slides is equal {pr.get_slide_count()}',
    ],
    model=model,
    response_model=StyleAgentResponse,
    debug_mode=True,
)

slide_id = 3

slide_schema = pr.get_json_schema()[slide_id]
slide_image_schema = pr.get_image_json(slide_id)
slide_text_schema = pr.get_text_frame_json(slide_id)


slide_images = []
for slide_id in range(1, pr.get_slide_count() + 1):
    slide_images.append(Image(content=pr.get_slide_image(slide_id), format='png'))

text_for_new_slide = dedent("""
РЕЗУЛЬТАТ ОТ ИСПОЛЬЗОВАНИЯ СИСТЕМЫ
1. Защита файлов от несанкционированного доступа
2. Снижение экономического ущерба
3. Ограниченный доступ к ключам
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
Вызовы и перспективы развития вычислительных мощностей
Рост цифровизации и внедрение искусственного интеллекта требуют постоянного увеличения производительности вычислительных систем.
Современные ограничения развития вычислительных мощностей связаны как с техническими (замедление роста тактовой частоты, минимальные размеры транзисторов), так и с рыночными факторами (дефицит чипов, зависимость от глобальных цепочек поставок)
Для преодоления этих вызовов активно развиваются новые технологии: квантовые вычисления, нейроморфные и тензорные процессоры, а также программно-определяемые решения (виртуализация, дезагрегация памяти)
В России и мире наблюдается тенденция к диверсификации аппаратных платформ и поиску альтернативных способов реализации вычислений, что позволяет повысить устойчивость и гибкость цифровой инфраструктуры
Инвестиции в развитие новых вычислительных архитектур и отечественных разработок становятся ключевым фактором технологической независимости и лидерства в цифровой экономике
""")

message = style_agent.run(
    # f"сформируй промт для генерации нового слайда для данного текста: '{text_for_new_slide} по данному шаблону: \nтекстовые блоки: {slide_text_schema}\nизображения: {slide_image_schema}'",
    f"сформируй промт для генерации нового слайда для данного текста: '{big2_text_for_new_slide}, вот тебе пример презентации куда тебе надо вставить слайд, попробуй скопировать стиль и найти самый подходящий слайд и на основе его создать новый {pr.get_json_schema()}",
    images=slide_images,
)

style_response: StyleAgentResponse = message.content

logger.debug(f'style response = {json.dumps(style_response.model_dump(), indent=4)}')

style_response.slide_index_to_add = 13

print(pr.get_slide_count())

if style_response.slide_index_to_add > pr.get_slide_count():
    pr.add_slide_at_position(pr.get_slide_count() + 1, background_color=style_response.background_color_rgb)

    style_response.slide_index_to_add = pr.get_slide_count()


for text_block in style_response.text_blocks:
    pr.create_text_shape(style_response.slide_index_to_add, text_block)

for image_block in style_response.image_blocks:
    logger.debug(
        f'create_image вызвана с параметрами: slide_id={style_response.slide_index_to_add}, opts={image_block}'
    )
    response = open_sync_client.images.generate(
        model='dall-e-2',
        prompt=image_block.prompt,
        n=1,
        size='512x512',
        response_format='b64_json',
    )

    b64_data = response.data[0].b64_json

    image_bytes = base64.b64decode(b64_data)

    pr.create_image_shape(
        style_response.slide_index_to_add,
        CreateImageFrameOpts(
            left=image_block.left,
            top=image_block.top,
            width=image_block.width,
            height=image_block.height,
            image=image_bytes,
        ),
    )
#
# slide_size = pr.get_slide_size_px()
# slide_count = pr.get_slide_count()
#
#
# def create_image(prompt: str, slide_id: int, opts: ImageFrameOpts) -> str:
#     """
#     Генерирует изображение с помощью DALL-E 2 на основе предоставленного запроса и размещает его на указанном слайде.
#
#     Args:
#         prompt (str): Текстовый запрос для генерации изображения.
#         slide_id (int): Идентификатор слайда, на котором будет размещено сгенерированное изображение.
#         opts (ImageFrameOpts): Параметры конфигурации для позиционирования и изменения размера сгенерированного изображения.
#             - left (float): Расстояние от левого края слайда.
#             - top (float): Расстояние от верхнего края слайда.
#             - width (float): Ширина изображения.
#             - height (float): Высота изображения.
#
#     Returns:
#         str: Сообщение о результате операции с подробным описанием созданной фигуры с изображением.
#     """
#     logger.debug(f'create_image вызвана с параметрами: prompt={prompt}, slide_id={slide_id}, opts={opts}')
#
#     response = open_sync_client.images.generate(
#         model='dall-e-2',
#         prompt=prompt,
#         n=1,
#         size='512x512',
#         response_format='b64_json',
#     )
#
#     b64_data = response.data[0].b64_json
#
#     image_bytes = base64.b64decode(b64_data)
#
#     return pr.create_image_shape(
#         slide_id,
#         CreateImageFrameOpts(
#             left=opts.left,
#             top=opts.top,
#             width=opts.width,
#             height=opts.height,
#             image=image_bytes,
#         ),
#     )
#
#
# head_agent, text_agent, slide_agent, image_agent = get_head_agent()
#
# text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape]
# slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]
# image_agent.tools = [create_image]
#
# text_agent.instructions[-1] = f'структура текстовых элементов: {pr.get_all_text_frame_json()}'
# image_agent.instructions[-1] = f'структура картинок в презентации: {pr.get_all_image_json()}'
# slide_agent.instructions[-1] = f'кол-во слайдов: {pr.get_slide_count()}'
#
# text_agent.instructions[-2] = (
#     f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
# )
# image_agent.instructions[-2] = (
#     f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
# )
# slide_agent.instructions[-2] = (
#     f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
# )
#
# head_agent.run(style_response)

logger.debug(f'metrics = {style_agent.run_response.metrics}')

pr.save('style_test.pptx')
