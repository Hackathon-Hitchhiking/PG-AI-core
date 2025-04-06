import os

from getpass import getpass

import httpx

from agno.agent import Agent
from agno.media import Image
from agno.models.openai import OpenAIChat
from dotenv import load_dotenv
from openai import AsyncOpenAI, OpenAI

from pptx_manager.main import PPTXManager
from pptx_manager.models import ImageFrameOpts, TextFrameOpts


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

model = OpenAIChat(id='gpt-4o-mini', client=open_sync_client, async_client=open_async_client)

test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path)

style_agent = Agent(
    name='Style Agent',
    instructions=[
        'Проанализируй предоставленные шаблонные слайды и определи:',
        '1.Расположение текстовых блоков (заголовки, подзаголовки, основные пункты).',
        '2.Стиль изображений (размер, выравнивание, соотношение сторон).',
        '3.Цветовую палитру и шрифты.'
        f'учитывай, что размер слайда равен {pr.get_slide_size_px()} в пикселях'
        f'параметры, которые можно использовать для текстового блока {TextFrameOpts.model_fields.keys()}'
        f'параметры, которые можно использовать для картинки {ImageFrameOpts.model_fields.keys()}',
    ],
    model=model,
)

slide_id = 1

slide = pr.get_all_image_json()[slide_id]

slide_image = pr.get_slide_image(slide_id)

style_agent.print_response(
    f"сформируй промт для генерации нового слайда по шаблону: {slide} для данного текста: 'Наша команда развивает и внедряет ИИ-модели (в т.ч. LLM) в ИТ-ландшафт г. Москвы (Мос.ру, Активный гражданин, Миллион призов, Город заданий), разрабатывая нового Цифрового ассистента для москвичей и помогая модераторам ранжировать контент, обрабатывать заявки и обращения москвичей, генерировать новости и т.п. '",
    images=[Image(content=slide_image, format='png')],
)
