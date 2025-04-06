import os

from typing import BinaryIO

from aiogram import Bot
from aiogram.types import BufferedInputFile, FSInputFile


class ResponseHandler:
    def __init__(self, bot: Bot):
        self.bot = bot

    async def send_text_response(self, user_id: int, text: str) -> None:
        """Отправить текстовый ответ пользователю"""
        await self.bot.send_message(chat_id=user_id, text=text)

    async def send_image_response(self, user_id: int, image_data: str | bytes | BinaryIO, caption: str = None) -> None:
        """
        Отправить изображение пользователю

        :param user_id: ID пользователя
        :param image_data: Путь к файлу, байты изображения или файловый объект
        :param caption: Подпись к изображению
        """
        if isinstance(image_data, str):
            if os.path.exists(image_data):
                photo = FSInputFile(image_data)
                await self.bot.send_photo(chat_id=user_id, photo=photo, caption=caption)
        elif isinstance(image_data, (bytes, BinaryIO)):
            photo = BufferedInputFile(
                image_data if isinstance(image_data, bytes) else image_data.read(), filename='image.jpg'
            )
            await self.bot.send_photo(chat_id=user_id, photo=photo, caption=caption)

    async def send_presentation(
        self, user_id: int, presentation_data: str | bytes | BinaryIO, filename: str = 'presentation.pptx'
    ) -> None:
        """
        Отправить презентацию пользователю

        :param user_id: ID пользователя
        :param presentation_data: Путь к файлу, байты презентации или файловый объект
        :param filename: Имя файла презентации
        """
        if isinstance(presentation_data, str):
            if os.path.exists(presentation_data):
                document = FSInputFile(presentation_data)
                await self.bot.send_document(
                    chat_id=user_id,
                    document=document,
                    caption=f'Готовая презентация: {os.path.basename(presentation_data)}',
                )
        elif isinstance(presentation_data, (bytes, BinaryIO)):
            document = BufferedInputFile(
                presentation_data if isinstance(presentation_data, bytes) else presentation_data.read(),
                filename=filename,
            )
            await self.bot.send_document(chat_id=user_id, document=document, caption=f'Готовая презентация: {filename}')
