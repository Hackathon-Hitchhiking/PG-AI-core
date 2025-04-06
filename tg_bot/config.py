import os

from dotenv import load_dotenv


load_dotenv()

API_TOKEN = os.getenv('BOT_TOKEN')
ADMIN_ID = os.getenv('ADMIN_ID')
TEMPLATES_DIR = os.path.join(os.getcwd(), 'downloads_template')
DOWNLOADS_DIR = os.path.join(os.getcwd(), 'downloads')
