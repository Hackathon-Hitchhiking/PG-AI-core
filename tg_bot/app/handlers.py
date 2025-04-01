import os

from aiogram import Bot, F, Router, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile

from pptx_manager.main import PPTXManager
from tests.test_agno import head_agent, image_agent, slide_agent, text_agent
from tg_bot.config import DOWNLOADS_DIR, TEMPLATES_DIR

from .keyboards import get_main_keyboard, get_template_keyboard
from .user_data import UserDataManager
from .utils import delete_template, ensure_user_dirs, get_user_templates


router = Router()


class TemplateStates(StatesGroup):
    waiting_for_template = State()
    waiting_for_presentation = State()
    waiting_for_query = State()


@router.message(Command('start'))
async def cmd_start(message: types.Message):
    await message.answer(
        'Привет! Это редактор презентаций. Я могу помочь тебе составить базовую презентацию на основе твоих запросов',
        reply_markup=get_main_keyboard(),
    )


@router.callback_query(F.data == 'upload_template')
async def upload_template(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer('Пожалуйста, отправьте файл шаблона в формате .pptx')
    await state.set_state(TemplateStates.waiting_for_template)


@router.message(TemplateStates.waiting_for_template, F.document)
async def handle_template_document(message: types.Message, state: FSMContext, bot: Bot):
    if not message.document.file_name.endswith('.pptx'):
        await message.answer('Пожалуйста, отправьте файл в формате .pptx')
        return

    user_dir = ensure_user_dirs(message.from_user.id)
    template_dir = os.path.join(TEMPLATES_DIR, str(message.from_user.id))
    file_path = os.path.join(template_dir, message.document.file_name)

    # Получаем file_id документа
    file_id = message.document.file_id

    # Получаем информацию о файле
    file = await bot.get_file(file_id)

    # Скачиваем файл
    await bot.download_file(file.file_path, file_path)

    await message.answer(f'Шаблон {message.document.file_name} успешно загружен!')
    await state.clear()


@router.callback_query(F.data == 'my_templates')
async def show_templates(callback: types.CallbackQuery):
    templates = get_user_templates(callback.from_user.id)
    if not templates:
        await callback.message.answer('У вас пока нет загруженных шаблонов.')
    else:
        for template in templates:
            await callback.message.answer(f'Шаблон: {template}', reply_markup=get_template_keyboard(template))


@router.callback_query(F.data.startswith('delete_'))
async def delete_template_handler(callback: types.CallbackQuery):
    template_name = callback.data.split('_', 1)[1]
    if delete_template(callback.from_user.id, template_name):
        await callback.message.answer(f'Шаблон {template_name} удален.')
    else:
        await callback.message.answer('Произошла ошибка при удалении шаблона.')


@router.callback_query(F.data.startswith('select_'))
async def select_template_handler(callback: types.CallbackQuery, state: FSMContext):
    template_name = callback.data.split('_', 1)[1]
    await state.update_data(selected_template=template_name)
    await callback.message.answer(f'Выбран шаблон: {template_name}')


@router.callback_query(F.data == 'upload_presentation')
async def upload_presentation(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer('Пожалуйста, отправьте файл презентации в формате .pptx для редактирования')
    await state.set_state(TemplateStates.waiting_for_presentation)


@router.message(TemplateStates.waiting_for_presentation, F.document)
async def handle_presentation_document(message: types.Message, state: FSMContext, bot: Bot):
    if not message.document.file_name.endswith('.pptx'):
        await message.answer('Пожалуйста, отправьте файл в формате .pptx')
        return

    user_dir = ensure_user_dirs(message.from_user.id)
    downloads_dir = os.path.join(DOWNLOADS_DIR, str(message.from_user.id))
    file_path = os.path.join(downloads_dir, message.document.file_name)

    # Получаем file_id документа
    file_id = message.document.file_id

    # Получаем информацию о файле
    file = await bot.get_file(file_id)

    # Скачиваем файл
    await bot.download_file(file.file_path, file_path)

    await state.set_data({'file_path': file_path})

    await message.answer(
        f'Файл {message.document.file_name} успешно загружен! Теперь опишите, что вы хотели бы изменить в презентации.'
    )
    await state.set_state(TemplateStates.waiting_for_query)


@router.message(TemplateStates.waiting_for_query, F.text)
async def handle_user_query(message: types.Message, state: FSMContext, bot: Bot):
    # Сохраняем запрос пользователя
    await UserDataManager.save_user_query(message.from_user.id, message.text, state)

    data = await state.get_data()

    file_path = data['file_path']

    pr = PPTXManager(file_path)

    text_agent.instructions[-1] = (f'структура текстовых элементов: {pr.get_all_text_frame_json()}',)
    image_agent.instructions[-1] = (f'структура картинок в презентации: {pr.get_all_image_json()}',)
    slide_agent.instructions[-1] = f'кол-во слайдов: {pr.get_slide_count()}'

    text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape]
    slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]

    head_agent.run(message.text)

    pr.save(file_path)

    document = FSInputFile(file_path)

    await message.reply_document(document, caption='Ваша призентаций')

    await message.answer(
        'Ваш запрос принят! Мы обрабатываем вашу презентацию. Вы можете отправить дополнительные комментарии или инструкции.'
    )


@router.callback_query(F.data == 'back_to_templates')
async def back_to_templates(callback: types.CallbackQuery):
    await show_templates(callback)
