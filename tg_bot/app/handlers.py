import base64
import os

from aiogram import Bot, F, Router, types
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile
from loguru import logger

from pptx_manager.main import PPTXManager
from pptx_manager.models import CreateImageFrameOpts, ImageFrameOpts
from tests.agno_manager import get_head_agent, open_sync_client
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

    # Сразу отправляем сообщение о принятии запроса
    processing_message = await message.answer(
        'Ваш запрос принят! Мы обрабатываем вашу презентацию. Это может занять некоторое время...'
    )

    # Получаем данные из состояния
    data = await state.get_data()
    file_path = data['file_path']

    try:
        # Отправляем индикатор набора текста, чтобы пользователь видел, что бот "работает"
        await bot.send_chat_action(message.chat.id, 'upload_document')

        pr = PPTXManager(file_path)

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

        text_agent.instructions[-1] = (f'структура текстовых элементов: {pr.get_all_text_frame_json()}',)
        image_agent.instructions[-1] = (f'структура картинок в презентации: {pr.get_all_image_json()}',)
        slide_agent.instructions[-1] = f'кол-во слайдов: {pr.get_slide_count()}'

        text_agent.tools = [pr.update_text_frame_shape, pr.create_text_shape, pr.delete_text_shape]
        slide_agent.tools = [pr.add_slide_at_position, pr.swap_slides]
        image_agent.tools = [create_image]

        # Периодически отправляем индикатор набора текста, чтобы пользователь видел, что бот "работает"
        await bot.send_chat_action(message.chat.id, 'upload_document')

        # Выполняем обработку запроса
        head_agent.run(message.text)

        # Сохраняем результат
        pr.save(file_path)

        # Отправляем обработанную презентацию
        document = FSInputFile(file_path)
        await message.reply_document(document, caption='Ваша презентация готова!')

        # Обновляем сообщение о статусе обработки
        await bot.edit_message_text(
            'Обработка завершена! Вы можете отправить дополнительные комментарии или инструкции для дальнейших изменений.',
            chat_id=message.chat.id,
            message_id=processing_message.message_id,
        )

    except Exception as e:
        logger.error(f'Ошибка при обработке презентации: {e}')
        await bot.edit_message_text(
            'Произошла ошибка при обработке вашей презентации. Пожалуйста, попробуйте еще раз или обратитесь в поддержку.',
            chat_id=message.chat.id,
            message_id=processing_message.message_id,
        )


@router.callback_query(F.data == 'back_to_templates')
async def back_to_templates(callback: types.CallbackQuery):
    await show_templates(callback)


@router.callback_query(F.data == 'coming_soon')
async def coming_soon_handler(callback: types.CallbackQuery):
    await callback.answer('Эта функция будет доступна в ближайшее время!', show_alert=True)
