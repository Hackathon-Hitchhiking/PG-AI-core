import base64
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

from pptx_manager.main import PPTXManager
from pptx_manager.models import CreateImageFrameOpts, ImageFrameOpts, TextFrameOpts
from tests.agno_manager import get_head_agent


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
        '3.Цветовую палитру и шрифты.',
        'Всегда копируй стиль слайда, который тебе передан, учитывай информацию, которую тебе передали, как текстом, так и картинкой',
        'Все цвета всегда указывай в RGB',
        'Учитывай, что ты генерируешь промт для другого агента, который будет как раз формировать данный слайд, не используй не какие дополнительные форматы',
        'Используй просто текст, не надо формировать не какие другие форматы',
        'Всегда указывай на какой слайд добавлять новые объекты',
        'Если это требуется ты должен использовать списки для формирование текстового блока',
        'Если это требуется ты должен использовать изображения',
        'При формирование изображения указывай там где оно должно разместиться и так же промт для генерации данного изображения',
        'Если пользователь хочет создать новый слайд, надо указывать, что нужно создать новый слайд и разместить там новые объекты',
        f'параметры, которые можно использовать для текстового блока {TextFrameOpts.model_fields.keys()}',
        f'параметры, которые нужно использовать для размещение всех объектов на слайде {ImageFrameOpts.model_fields.keys()}, используй их вместе со всеми блоками',
        f'учитывай, что размер слайда равен {pr.get_slide_size_px()} в пикселях, когда будешь размещать объекты',
        f'учитывай, что количество слайдов равно {pr.get_slide_count()}',
    ],
    model=model,
)

slide_id = 9

slide = pr.get_all_image_json()[slide_id]

slide_image = pr.get_slide_image(slide_id)

text_for_new_slide = dedent("""
Видеоаналитика – эффективный и перспективный инструмент для большинства отраслей городского управления. Задача – расширить ее применение. 
В сфере безопасности аналитика позволит фиксировать нестандартное поведения отдельных людей, аномальные скопления толп, продолжит улучшать процесс поиска правонарушителей и пропавших граждан. Для расследования и предотвращения преступлений будут развиваться алгоритмы анализа больших данных видеонаблюдения.
Новые алгоритмы для анализа объектов городской инфраструктуры помогут выявлять еще больше недочетов в ЖКХ и сфере землепользования: следить за содержанием объектов, территорий, а также мониторить работы по благоустройству и строительству. 
Используя данные от ИИ, который будет анализировать пути движения пешеходов и пользователей СИМ, можно будет формировать оптимальные варианты для организации пешеходных переходов и других объектов улично-дорожной сети. 
Компьютерное зрение найдет применение и в сфере массового обслуживания. Это мониторинг очередей в МФЦ, объектах здравоохранения и соцсферы.
Планируется и развитие инструментов дополнительного анализа видеоданных. Пользователи ЕЦХД смогут переходить в трехмерное видеопространство, как способ более эффективного получения информации о текущей ситуации или навигации по архивным данным в прошлом «внутри» цифрового двойника Москвы. Конвергентная умная разметка видеополя (фиксация на изображении различных объектов и их состояний) позволит быстро находить изображения и увеличит срок хранения полезной информации в архиве.
""")

message = style_agent.run(
    f"сформируй промт для генерации нового слайда по данному шаблону: {slide} для данного текста: '{text_for_new_slide}'",
    images=[Image(content=slide_image, format='png')],
)

style_response = message.content

logger.debug(f'style response = {style_response}')

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


head_agent, text_agent, slide_agent, image_agent = get_head_agent()

text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape]
slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]
image_agent.tools = [create_image]

text_agent.instructions[-1] = f'структура текстовых элементов: {pr.get_all_text_frame_json()}'
image_agent.instructions[-1] = f'структура картинок в презентации: {pr.get_all_image_json()}'
slide_agent.instructions[-1] = f'кол-во слайдов: {pr.get_slide_count()}'

text_agent.instructions[-2] = (
    f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
)
image_agent.instructions[-2] = (
    f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
)
slide_agent.instructions[-2] = (
    f'Размер слайдов в пикселях {slide_size}, используй координаты, чтобы вставлять объекты.',
)

head_agent.run(style_response)

pr.save('style_test.pptx')
