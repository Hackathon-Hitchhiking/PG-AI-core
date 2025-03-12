class TextAgent:
    """
    Агент для генерации текстового содержимого слайдов.
    Поддерживает как локальные модели, так и внешние API.
    """
    
    def __init__(self, model_type="api", model_config=None):
        """
        Инициализация с типом модели и конфигурацией.
        model_type: "api" или "local"
        """
        pass
    
    def generate_title(self, slide_context):
        """Генерация заголовка слайда."""
        pass
    
    def generate_content(self, slide_context, placeholder_info):
        """Генерация основного текста для указанного плейсхолдера."""
        pass
    
    def format_text(self, text, placeholder_constraints):
        """Форматирование текста под ограничения плейсхолдера."""
        pass
