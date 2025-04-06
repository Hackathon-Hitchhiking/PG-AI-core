from aiogram.types import InlineKeyboardButton
from aiogram.utils.keyboard import InlineKeyboardBuilder


def get_main_keyboard():
    builder = InlineKeyboardBuilder()
    builder.add(InlineKeyboardButton(text='Загрузить шаблон', callback_data='upload_template'))
    builder.add(InlineKeyboardButton(text='Мои шаблоны', callback_data='my_templates'))
    builder.add(InlineKeyboardButton(text='Загрузить презентацию', callback_data='upload_presentation'))
    builder.add(InlineKeyboardButton(text='Причесать загруженную презентацию', callback_data='coming_soon'))
    builder.adjust(2, 2)
    return builder.as_markup()


def get_template_keyboard(template_name):
    builder = InlineKeyboardBuilder()
    builder.add(InlineKeyboardButton(text='Выбрать', callback_data=f'select_{template_name}'))
    builder.add(InlineKeyboardButton(text='Удалить', callback_data=f'delete_{template_name}'))
    builder.add(InlineKeyboardButton(text='Назад', callback_data='back_to_templates'))
    return builder.as_markup()
