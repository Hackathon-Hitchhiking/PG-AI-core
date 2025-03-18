# config.py
CONFIG = {
    # Модели и API
    'text_model': {
        'type': 'api',  # 'api' или 'local'
        'api_endpoint': 'https://api.example.com/v1/completions',
        'api_key': 'YOUR_API_KEY',
        'model_name': 'rugpt-3.5-turbo',
        'local_path': 'models/rugpt-3.5-turbo',
        'temperature': 0.7
    },

    'image_model': {
        'type': 'api',  # 'api' или 'local'
        'api_endpoint': 'https://api.example.com/v1/images/generations',
        'api_key': 'YOUR_API_KEY',
        'model_name': 'kandinsky-3',
        'local_path': 'models/kandinsky-3',
        'guidance_scale': 7.5
    },

    # Пути к ресурсам
    'templates_dir': 'resources/templates',
    'output_dir': 'output/presentations',
    'prompt_templates': 'resources/prompts.json',

    # Параметры генерации
    'default_slide_count': 10,
    'max_retries': 3,
    'timeout': 60
}
