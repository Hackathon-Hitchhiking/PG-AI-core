from abc import abstractmethod, ABC

class ImageGeneratorInterface(ABC):
    """
    Интерфейс для генераторов изображений.
    Определяет методы для генерации различных типов визуального контента.
    """
    
    @abstractmethod
    def generate_illustration(self, prompt, specifications):
        """Генерирует иллюстрацию по текстовому описанию."""
        pass
    
    @abstractmethod
    def generate_background(self, theme, color_scheme, specifications):
        """Генерирует фоновое изображение для слайда."""
        pass
    
    @abstractmethod
    def generate_diagram(self, data, diagram_type, specifications):
        """Генерирует диаграмму или график на основе данных."""
        pass
    
    @abstractmethod
    def adapt_image_to_placeholder(self, image, placeholder_constraints):
        """Адаптирует изображение под ограничения плейсхолдера."""
        pass
    
    @abstractmethod
    def ensure_visual_consistency(self, image, presentation_style):
        """Обеспечивает визуальную согласованность с другими элементами."""
        pass
