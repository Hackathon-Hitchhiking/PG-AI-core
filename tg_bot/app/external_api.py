from typing import BinaryIO

from aiogram import Bot
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.base import BaseStorage

from .response_handler import ResponseHandler
from .user_data import UserDataManager


class ExternalAPI:
    def __init__(self, bot: Bot, storage: BaseStorage):
        self.bot = bot
        self.storage = storage
        self.response_handler = ResponseHandler(bot)

    async def get_user_state(self, user_id: int) -> FSMContext:
        """Получить объект состояния пользователя"""
        return FSMContext(storage=self.storage, key=f'user:{user_id}')

    async def get_selected_template(self, user_id: int) -> str | None:
        """Получить выбранный пользователем шаблон"""
        state = await self.get_user_state(user_id)
        return await UserDataManager.get_selected_template(user_id, state)

    async def get_uploaded_presentation(self, user_id: int) -> str | None:
        """Получить загруженную пользователем презентацию"""
        return await UserDataManager.get_uploaded_presentation(user_id)

    async def get_user_queries(self, user_id: int) -> list[str]:
        """Получить все текстовые запросы пользователя"""
        state = await self.get_user_state(user_id)
        return await UserDataManager.get_user_queries(user_id, state)

    async def send_text_response(self, user_id: int, text: str) -> None:
        """Отправить текстовый ответ пользователю"""
        await self.response_handler.send_text_response(user_id, text)

    async def send_image_response(self, user_id: int, image_data: str | bytes | BinaryIO, caption: str = None) -> None:
        """Отправить изображение пользователю"""
        await self.response_handler.send_image_response(user_id, image_data, caption)

    async def send_presentation(
        self, user_id: int, presentation_data: str | bytes | BinaryIO, filename: str = 'presentation.pptx'
    ) -> None:
        """Отправить презентацию пользователю"""
        await self.response_handler.send_presentation(user_id, presentation_data, filename)

    async def clear_user_queries(self, user_id: int) -> None:
        """Очистить историю запросов пользователя"""
        state = await self.get_user_state(user_id)
        await UserDataManager.clear_user_queries(user_id, state)
