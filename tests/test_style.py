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
from pptx_manager.models import ImageFrameOpts
from tests.agno_manager import get_pptx_agent
from tests.constants import (
    CREATOR_PPTX_AGENT_INSTRUCTIONS,
    FINALIZER_PPTX_AGENT_INSTRUCTIONS,
    STYLE_AGENT_CORE_INSTRUCTIONS_V2,
)


logger.add('test.log', rotation='100 MB', encoding='utf-8')
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


style_agent_model = OpenAIChat(id='gpt-4.1', client=open_sync_client, async_client=open_async_client)

test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path, False)

style_agent = Agent(
    name='Style Agent',
    instructions=STYLE_AGENT_CORE_INSTRUCTIONS_V2
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


# slide_images = []
# for slide_id in range(1, pr.get_slide_count() + 1):
#     slide_images.append(Image(content=pr.get_slide_image(slide_id), format='png'))

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

user_request_for_final_test_1 = dedent("""
Добавь новый слайд в конец
Отличительные особенности:
возможность интеграции с ЕСИА для авторизации и соблюдения мер ИБ и 152-ФЗ;
гибкая настройка курсов для разных потребностей и уровня подготовки;
возможность интеграции с внешними системами для обогащения данными и для встраивания в различные бизнес-процессы, например интеграция с системами мониторинга;
безопасность – решение состоит в Реестре Отечественного ПО, стек соответствует требованиям ИБ.""")

user_request_for_final_test_2 = dedent("""
Добавь новый слайд в конец
Система управления обучением Вектор
Начало карьеры:
тестирование для профориентации
вводные курсы для ознакомления с профессией
онбординг новых сотрудников
самообразование
Развитие карьеры: 
дополнительное профессиональное образование
целевое обучение
аттестация на соответствие квалификации
удержание сотрудника с помощью корпоративных программ обучения (повышается ценность образования)
генерация защищенного сертификата об обучении 
отслеживание и оценка прогресса обучения, формирование кадрового резерва
Карьерные траектории: 
анализ навыков и дополнительная профориентация
дообучение для смежных отраслей
авторские курсы и передача опыта
""")

tasks = pr.get_tasks_from_slide()

# for slide_id, task in tasks.items():
#     message = style_agent.run(
#         f'Проанализируй этот пользовательский запрос и эталонную презентацию. Сгенерируй подробные инструкции для HeadAgent для слайда {slide_id}: {task}',
#         # images=slide_images,
#     )

message = style_agent.run(
    f'Проанализируй этот пользовательский запрос и эталонную презентацию. Сгенерируй подробные инструкции для PPTXAgent: {user_request_for_final_test_2}',
    # images=slide_images,
)

style_output: StyleOutput = message.content

logger.debug(f'StyleAgent instructions:\n{json.dumps(style_output.model_dump(), indent=4, ensure_ascii=False)}')

# head_agent, text_agent, slide_agent, image_agent, figure_agent = get_head_agent(pr)
#
# head_agent.run(style_output.instructions)
#
#
# logger.debug(f'head agent metrics = {head_agent.run_response.metrics if head_agent.run_response else 0}')
# logger.debug(f'image agent metrics = {image_agent.run_response.metrics if image_agent.run_response else 0}')
# logger.debug(f'text agent metrics = {text_agent.run_response.metrics if text_agent.run_response else 0}')
# logger.debug(f'slide agent metrics = {slide_agent.run_response.metrics if slide_agent.run_response else 0}')
# logger.debug(f'figure agent metrics = {figure_agent.run_response.metrics if figure_agent.run_response else 0}')

creator_pptx_agent = get_pptx_agent(pr, CREATOR_PPTX_AGENT_INSTRUCTIONS)

creator_pptx_agent.run('\n'.join(style_output.instructions))

pr.save('style_test_before_finalizer.pptx')

finalizer_pptx_agent = get_pptx_agent(pr, FINALIZER_PPTX_AGENT_INSTRUCTIONS)

pr.parse_slide_as_images()

with open('test.png', 'wb') as f:
    f.write(pr.get_slide_image(8))

finalizer_pptx_agent.run(
    f'Провалидируй данную презентацию, тебе подан сейчас только 14-ый слайд, провалидируй только его: {pr.get_json_schema()[8]}',
    images=[Image(content=pr.get_slide_image(8), format='png')],
)

logger.debug(f'style agent metrics = {style_agent.run_response.metrics}')
logger.debug(f'creator_pptx_agent metrics = {creator_pptx_agent.run_response.metrics}')
logger.debug(f'finalizer_pptx_agent metrics = {finalizer_pptx_agent.run_response.metrics}')

pr.save('style_test.pptx')
