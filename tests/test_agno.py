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
    tools=[pr.add_slide_at_position],
    model=model,
    show_tool_calls=True,
    debug_mode=True,
)

image_agent = Agent(
    name='Image Agent',
    instructions=[
        'You are a specialized agent responsible for managing and optimizing visual content within PowerPoint presentations.',
        'Your core responsibilities include:',
        '- Handling image insertions, replacements, and modifications',
        '- Ensuring image quality and resolution standards',
        '- Maintaining proper aspect ratios and positioning',
        '- Managing image sizing and formatting',
        '- Working with slide_id and shape_id to precisely locate and modify images',
        'Always confirm visual changes and provide feedback on completed operations.',
        "Ensure all image modifications align with presentation's overall design and purpose.",
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
    instructions=[
        'You are a Lead Presentation Manager responsible for orchestrating changes in PowerPoint presentations.',
        'Key responsibilities:',
        '- Coordinate with specialized agents to implement user-requested modifications',
        '- Ensure accurate interpretation of user requirements',
        '- Validate changes using slide_id (slide number) and shape_id (element identifier)',
        '- Maintain presentation consistency and professional quality',
        '- Confirm successful completion of all modifications',
        'Process:',
        '1. Analyze user request',
        '2. Route tasks to appropriate specialized agents',
        '3. Verify execution',
        '4. Report outcomes clearly',
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
