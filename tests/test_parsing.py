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


STYLE_AGENT_CORE_INSTRUCTIONS = [
    # Analyze the reference presentation for design language, layout, and content structure.
    'Analyze the reference presentation to extract its visual style, layout patterns, and content organization.',
    'Identify color schemes, typography, spacing, and recurring design elements.',
    'Determine how new content should be integrated to match the existing style.',
    # Generate actionable instructions for HeadAgent.
    'For the given user request, generate a detailed list of instructions describing:',
    '- Which slides to add or modify.',
    '- Where to place each content element (text, images, etc.) with coordinates and sizes.',
    '- What content to include (text, images, background, etc.).',
    '- Any specific style requirements (colors, fonts, effects).',
    'Output only a list of clear, step-by-step instructions for HeadAgent to execute.',
    'Do not perform any slide or image creation yourself.',
    'Ensure instructions are unambiguous and cover all necessary details for implementation.',
]


model = OpenAIChat(id='gpt-4o', client=open_sync_client, async_client=open_async_client)

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
    model=model,
    response_model=StyleOutput,
    debug_mode=True,
)


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

user_request = dedent("""
Вызовы и перспективы развития вычислительных мощностей
Рост цифровизации и внедрение искусственного интеллекта требуют постоянного увеличения производительности вычислительных систем.
Современные ограничения развития вычислительных мощностей связаны как с техническими (замедление роста тактовой частоты, минимальные размеры транзисторов), так и с рыночными факторами (дефицит чипов, зависимость от глобальных цепочек поставок)
Для преодоления этих вызовов активно развиваются новые технологии: квантовые вычисления, нейроморфные и тензорные процессоры, а также программно-определяемые решения (виртуализация, дезагрегация памяти)
В России и мире наблюдается тенденция к диверсификации аппаратных платформ и поиску альтернативных способов реализации вычислений, что позволяет повысить устойчивость и гибкость цифровой инфраструктуры
Инвестиции в развитие новых вычислительных архитектур и отечественных разработок становятся ключевым фактором технологической независимости и лидерства в цифровой экономике
""")

message = style_agent.run(
    f'Проанализируй этот пользовательский запрос и эталонную презентацию. Сгенерируй подробные инструкции для HeadAgent по добавлению нового слайда и размещению всех элементов: {user_request}',
    images=slide_images,
)

style_output: StyleOutput = message.content

logger.debug(f'StyleAgent instructions:\n{json.dumps(style_output.model_dump(), indent=4, ensure_ascii=False)}')

head_agent, text_agent, slide_agent, image_agent = get_head_agent()

text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape]
slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]
image_agent.tools = [create_image]

text_agent.instructions[-1] = f'структура текстовых элементов: {pr.get_all_text_frame_json()}'
image_agent.instructions[-1] = f'структура картинок в презентации: {pr.get_all_image_json()}'
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

head_agent.run(style_output.instructions)

logger.debug(f'metrics = {style_agent.run_response.metrics}')

pr.save('style_test.pptx')
