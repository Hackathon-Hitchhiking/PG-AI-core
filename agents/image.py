class ImageAgent:
    """
    Агент для генерации изображений для слайдов.
    Поддерживает как локальные модели, так и внешние API.
    """
    
    def __init__(self, model_type="api", model_config=None):
        """
        Инициализация с типом модели и конфигурацией.
        model_type: "api" или "local"
        """
        pass
    
    def generate_image(self, prompt, size, style_params=None):
        """Генерация изображения по текстовому описанию с указанными параметрами."""
        pass
    
    def adjust_image(self, image, target_size, crop=True):
        """Корректировка размера и пропорций изображения под плейсхолдер."""
        pass
    
    def apply_style(self, image, style_params):
        """Применение стилевых эффектов к изображению."""
        pass
