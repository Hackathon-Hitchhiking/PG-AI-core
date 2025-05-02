import base64
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
from pptx_manager.models import CreateImageFrameOpts, ImageFrameOpts


load_dotenv()

if not os.environ.get('OPENAI_API_KEY'):
    os.environ['OPENAI_API_KEY'] = getpass.getpass('Enter API key for OpenAI: ')


def str2bool(value: str) -> bool:
    return str(value).lower() in ('true', '1', 'yes', 'y', 'on')


if str2bool(os.environ.get('USE_PROXY_URLS', 'True')):
    http_async_client = httpx.AsyncClient(proxy='socks5://127.0.0.1:12334')
    http_sync_client = httpx.Client(proxy='socks5://127.0.0.1:12334')
else:
    http_async_client = None
    http_sync_client = None

open_async_client = AsyncOpenAI(
    http_client=http_async_client,
)

open_sync_client = OpenAI(
    http_client=http_sync_client,
)

model_41 = OpenAIChat(id='gpt-4.1', client=open_sync_client, async_client=open_async_client)
model_4o = OpenAIChat(id='gpt-4o', client=open_sync_client, async_client=open_async_client)


def get_head_agent(pr: PPTXManager):
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

        # Append minimalistic style and no text requirements to the prompt
        enhanced_prompt = f'{prompt}. The image should be in a minimalistic style and contain no text.'

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

    text_agent = Agent(
        name='Text Agent',
        instructions=[
            # Роль
            'Эксперт по работе с текстом в PowerPoint.',
            'Выполняй только свою зону ответственности; если шаг вне компетенции — пропусти его.',
            # Базовые правила
            'Всегда указывай slide_id и shape_id.',
            'Сохраняй исходное форматирование (шрифт, размер, цвет).',
            'Не меняй макет слайда без прямого запроса.',
            # Поддерживаемые действия
            '- update_text_frame_shape — заменить содержимое.',
            '- create_text_shape — добавить новый блок текста.',
            # Проверки
            'Перед действием убеждайся, что slide_id/shape_id существуют.',
            # Автодополнение
            'Если нужного текста нет — добавь на основе контекста.',
            # Отчёт
            'В конце дай краткий отчёт: что, где, как изменено.',
            # Служебная информация
            f'Размер слайда (px): {slide_size}',
            f'JSON-схема текстовых элементов: {pr.get_all_text_frame_json()}',
        ],
        tools=[pr.update_text_frame_shape, pr.create_text_shape],  # pr.delete_text_shape
        model=model_4o,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    slide_agent = Agent(
        name='Slide Agent',
        role='Presentation Slide Content Specialist',
        instructions=[
            'Специалист по управлению количеством слайдов.',
            'Выполняй только свою зону ответственности; если шаг вне компетенции — пропусти его.',
            # Добавление
            'При добавлении без позиции — вставляй в конец.',
            '"в начало" → позиция 1, "в конец" → последняя, "в середину" → (n//2)+1.',
            # Правила
            'Всегда добавляй **ровно один** слайд.',
            'Не делай предположений; используй только входные данные.',
            'Соблюдай размеры слайда и координатную сетку.',
            'Никогда не спрашивай подтверждения у пользователя.',
            'У тебя нет саб агентов не надо их вызывать',
            'Не надо вызывать функции о которых тебе не говорили',
            # Отчёт
            'Сообщай: создан слайд #{new_id}.',
            # Служебная информация
            f'Размер слайда (px): {slide_size}',
            f'Текущее количество слайдов: {slide_count}',
        ],
        tools=[pr.add_slide_at_position],  # pr.delete_slide, pr.swap_slides
        model=model_41,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    image_agent = Agent(
        name='Image Agent',
        instructions=[
            'Визуальный инженер: вставка и замена изображений.',
            'Выполняй только свою зону ответственности; если шаг вне компетенции — пропусти его.',
            # Основные задачи
            'Работай только по slide_id/shape_id.',
            'Авто-оптимизируй размер и позицию по сетке 0.1 см.',
            'Не генерируй несколько изображений с одним и тем же prompt.',
            # Процесс
            '1. Проверка существования slide_id/shape_id и формата файла.',
            '2. Выполнение (insert/replace) с оптимизацией.',
            '3. Верификация: хэш, метаданные, слои.',
            # Отчёт
            'Отчитывайся: действие, slide_id, shape_id, размер.',
            # Служебная информация
            'Размер слайда (px): {slide_size}',
            f'JSON-схема изображений: {pr.get_all_image_json()}',
        ],
        tools=[create_image, pr.copy_image_shape],  # pr.delete_image_shape
        model=model_41,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    figure_agent = Agent(
        name='Figure Agent',
        instructions=[
            'Эксперт по геометрическим фигурам.',
            'Выполняй только свою зону ответственности; если шаг вне компетенции — пропусти его.',
            # Действия
            '- add_figure_shape',
            '- update_shape_color / position / transparency',
            '- set_shape_rounding',
            '- copy_figure_shape',
            # Правила
            'Всегда проверяй существование фигуры перед изменением.',
            'Координаты и размеры — в пикселях.',
            'Цвет — RGB 0-255; прозрачность 0.0-1.0; скругление 0.0-1.0.',
            'Не трогай неупомянутые фигуры.',
            'При недостатке параметров — используй разумные значения.',
            # Создание нового слайда
            'Если требуется новый слайд — сначала через Slide Agent.',
            # Отчёт
            'Сообщай: действие, slide_id, shape_id/тип, параметры.',
            # Служебная информация
            f'Размер слайда (px): {slide_size}',
            f'JSON-схема фигур: {pr.get_all_figure_frame_json()}',
        ],
        tools=[
            pr.update_shape_color,
            pr.update_shape_position,
            pr.update_shape_transparency,
            pr.set_shape_rounding,
            pr.add_figure_shape,
            pr.copy_figure_shape,
            #    pr.delete_figure_shape,
        ],
        model=model_41,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    head_agent = Team(
        mode='coordinate',
        members=[text_agent, slide_agent, image_agent, figure_agent],
        model=model_41,
        instructions=[
            'Главный координатор правок презентации.',
            # Алгоритм
            '1. Разбирай запрос пользователя.',
            "2. Делегируй задачи Text/Image/Figure/Slide-Agent'ам.",
            '3. Собирай и проверяй результаты (позиционирование, стиль, целостность).',
            '4. При ошибке перезапусти действие (не более 2 раз).',
            # Жёсткие правила
            'Не выдумывай новых агентов, используй только заданные.',
            'Не запрашивай подтверждений у пользователя.',
            'Всегда передавай slide_id в под-агенты.',
            'При создании слайда: сначала Slide Agent, затем остальные.',
        ],
        show_tool_calls=True,
        monitoring=False,
        telemetry=False,
        show_members_responses=True,
        debug_mode=True,
    )

    return head_agent, text_agent, slide_agent, image_agent, figure_agent


def get_pptx_agent(pr: PPTXManager, instructions: list[str]):
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

        # Append minimalistic style and no text requirements to the prompt
        enhanced_prompt = f'{prompt}. The image should be in a minimalistic style and contain no text.'

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

    pptx_agent = Agent(
        name='PPTX Agent',
        instructions=instructions
        + [
            f'Размер слайда (px): {slide_size}',
            f'Кол-во слайдов: {slide_count}',
        ],
        tools=[
            pr.update_shape_color,
            pr.update_shape_position,
            pr.update_shape_transparency,
            pr.set_shape_rounding,
            pr.add_figure_shape,
            pr.copy_figure_shape,
            create_image,
            pr.copy_image_shape,
            pr.add_slide_at_position,
            pr.update_text_frame_shape,
            pr.create_text_shape,
            #    pr.delete_figure_shape,
        ],
        model=model_4o,
        monitoring=False,
        telemetry=False,
        show_tool_calls=True,
        debug_mode=True,
    )

    return pptx_agent
