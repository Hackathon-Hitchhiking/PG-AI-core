import os

from aiogram.fsm.context import FSMContext

from tg_bot.config import DOWNLOADS_DIR, TEMPLATES_DIR


class UserDataManager:
    @staticmethod
    async def get_selected_template(user_id: int, state: FSMContext) -> str | None:
        """Получить путь к выбранному пользователем шаблону"""
        data = await state.get_data()
        selected_template = data.get('selected_template')

        if not selected_template:
            return None

        template_path = os.path.join(TEMPLATES_DIR, str(user_id), selected_template)
        if os.path.exists(template_path):
            return template_path
        return None

    @staticmethod
    async def get_uploaded_presentation(user_id: int) -> str | None:
        """Получить путь к последней загруженной пользователем презентации"""
        user_dir = os.path.join(DOWNLOADS_DIR, str(user_id))
        if not os.path.exists(user_dir):
            return None

        files = [f for f in os.listdir(user_dir) if f.endswith('.pptx')]
        if not files:
            return None

        latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(user_dir, f)))
        return os.path.join(user_dir, latest_file)

    @staticmethod
    async def save_user_query(user_id: int, query: str, state: FSMContext) -> None:
        """Сохранить текстовый запрос пользователя"""
        data = await state.get_data()
        queries = data.get('user_queries', [])
        queries.append(query)
        await state.update_data(user_queries=queries)

    @staticmethod
    async def get_user_queries(user_id: int, state: FSMContext) -> list[str]:
        """Получить все текстовые запросы пользователя"""
        data = await state.get_data()
        return data.get('user_queries', [])

    @staticmethod
    async def clear_user_queries(user_id: int, state: FSMContext) -> None:
        """Очистить историю запросов пользователя"""
        await state.update_data(user_queries=[])
