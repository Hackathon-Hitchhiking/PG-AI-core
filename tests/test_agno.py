import getpass
import os

import httpx

from agno.agent import Agent
from agno.models.openai import OpenAIChat
from agno.team import Team
from dotenv import load_dotenv
from loguru import logger
from openai import AsyncOpenAI, OpenAI

from pptx_manager.main import PPTXManager


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

test_pres_path = os.environ.get('TEST_PRES_PATH')
pr = PPTXManager(test_pres_path)

text_json = pr.get_all_text_frame_json()

# pr_json = pr.get_all_text_frame_json()

model = OpenAIChat(id='gpt-4o-mini', client=open_sync_client, async_client=open_async_client)

text_agent = Agent(
    name='Text Agent',
    instructions=[
        'Вы — эксперт по управлению текстом в PowerPoint, использующий slide_id (номер слайда) и shape_id (идентификатор текстового элемента) для точного определения местоположения элементов.',
        'Перед выполнением операции всегда проверяйте корректность slide_id и shape_id.',
        'Изменяйте текст с соблюдением следующих требований:',
        '- Сохраняйте исходное форматирование (шрифт, размер, цвет).',
        '- Поддерживайте структурную целостность слайда.',
        '- Учитывайте заданные параметры шрифта, размера и других стилей.',
        'После выполнения операции подтверждайте изменения и предоставляйте отчет о проделанной работе.',
        'Не изменяйте макет слайда или структуру презентации без явного запроса.',
        'Работайте только с указанными элементами. Не выполняйте предположений относительно контекста или содержимого.',
        f'структура текстовых элементов: {text_json}',
    ],
    tools=[pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape],
    model=model,
    show_tool_calls=True,
    markdown=True,
    debug_mode=True,
)

slide_agent = Agent(
    name='Slide Agent',
    role='Presentation Slide Content Specialist',
    instructions=[
        'Вы — эксперт по управлению количеством слайдов в PowerPoint.',
        'Работайте только с указанными элементами. Не выполняйте предположений относительно контекста или содержимого.',
    ],
    tools=[pr.add_slide_at_position, pr.swap_slides],
    model=model,
    show_tool_calls=True,
    debug_mode=True,
)

image_agent = Agent(
    name='Image Agent',
    instructions=[
        "Вы Визуальный Инженер презентаций: эксперт по работе с изображениями в PowerPoint",
        " Основные задачи:",
        "• Точная вставка/замена изображений (только по slide_id и shape_id)",
        "• Автокоррекция параметров",
        "• Позиционирование с привязкой к сетке (шаг 0.1 см)",
        "◉ Протокол работы:",
        "1. Получить задание: [Слайд][Элемент][Действие][Параметры]",
        "2. Проверить:",
        "   - Существование slide_id/shape_id",
        "   - Совместимость формата файла",
        "   - Соответствие бренд-буку (RGB/CMYK)",
        "3. Выполнить с автоматической оптимизацией:",
        "   - Ресайз через Lanczos3",
        "   - Цветокоррекция (гамма 2.2)",
        "   - Сжатие без потерь (quality=92)",
        "4. Верификация:",
        "   - Сравнение хэша изображения",
        "   - Проверка метаданных",
        "   - Контроль слоев",
    ],
    tools=[],
    model=model,
    show_tool_calls=True,
    debug_mode=True,
)

head_agent = Team(
    mode='coordinate',
    members=[text_agent, slide_agent],
    model=model,
    instructions = [
        "Вы главный по управлению презентациями: координируешь агентов для правок в PowerPoint",
        "Обязанности:",
        "• Распределяй задачи между агентами (текст, графика, макеты)",
        "• Контролируй выполнение по slide_id (номер слайда) и shape_id (ID элемента)",
        "• Гарантируй сохранение стиля и структуры презентации",
        "Этапы работы:",
        "1. Разбери запрос → определи нужных агентов",
        "2. Дай точные указания (слайд/элемент/действие)",
        "3. Проверь результат:",
        "   - Корректность позиционирования",
        "   - Соответствие стилю",
        "   - Целостность презентации",
        "При ошибках:",
        "• Повтори попытку (макс. 2 раза)",
        "• Подключи другого агента",
        "• Сообщи об проблеме",
    
    ],
    memory=None,
    context=None,
    show_tool_calls=True,
    show_members_responses=True,
    debug_mode=True,
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

    response = head_agent.run(f'{user_input}')
    print(response.content)
    logger.debug(f'formated_tool_calls = {response.formatted_tool_calls}')
    logger.debug(f'tools = {[response.tools for response in response.member_responses if response.tools is not None]}')
    # Save the updated presentation
    pr.save(f'test_{num}.pptx')
    logger.debug(f'Презентация сохранена как test_{num}.pptx')
    num += 1
