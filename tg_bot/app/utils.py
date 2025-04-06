import os

from tg_bot.config import DOWNLOADS_DIR, TEMPLATES_DIR


def get_user_templates(user_id):
    user_dir = os.path.join(TEMPLATES_DIR, str(user_id))
    if not os.path.exists(user_dir):
        return []
    return [f for f in os.listdir(user_dir) if f.endswith('.pptx')]


def delete_template(user_id, template_name):
    template_path = os.path.join(TEMPLATES_DIR, str(user_id), template_name)
    if os.path.exists(template_path):
        os.remove(template_path)
        return True
    return False


def ensure_user_dirs(user_id):
    for dir_path in [TEMPLATES_DIR, DOWNLOADS_DIR]:
        user_dir = os.path.join(dir_path, str(user_id))
        os.makedirs(user_dir, exist_ok=True)
    return user_dir
