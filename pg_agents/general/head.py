# agents/general/head.py
from agents import Agent

from . import prompts
from .chart import ChartAgent
from .image import ImageAgent
from .slide import SlideAgent
from .table import TableAgent
from .text import TextAgent


class HeadAgent(Agent):
    def __init__(self):
        super().__init__(
            name='HeadController',
            instructions=prompts.HEAD_PROMPT,
            tools=[
                TextAgent().as_tool(
                    tool_name='text_manager',
                    tool_description='Управляет текстовыми блоками: создание, редактирование, стилизация текста на слайдах.',
                ),
                SlideAgent().as_tool(
                    tool_name='slide_manager',
                    tool_description='Управляет структурой презентации: добавление, удаление и изменение порядка слайдов.',
                ),
                TableAgent().as_tool(
                    tool_name='table_manager',
                    tool_description='Управляет таблицами: создание, редактирование, форматирование и заполнение данными.',
                ),
                ImageAgent().as_tool(
                    tool_name='image_manager',
                    tool_description='Управляет изображениями: вставка, редактирование, изменение размеров и позиционирование изображений на слайдах.',
                ),
                ChartAgent().as_tool(
                    tool_name='chart_manager',
                    tool_description='Управляет графиками и диаграммами: создание, редактирование, настройка параметров и обновление данных для визуализации.',
                ),
            ],
            model='gpt-4-turbo',
        )
