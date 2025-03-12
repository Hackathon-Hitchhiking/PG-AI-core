from abc import abstractmethod, ABC


class TextGeneratorInterface(ABC):
    """
    Интерфейс для генераторов текста.
    Определяет методы для генерации различных типов текстового контента.
    """
    
    @abstractmethod
    def generate_title(self, context, specifications):
        """Генерирует заголовок слайда."""
        pass
    
    @abstractmethod
    def generate_body_text(self, context, specifications):
        """Генерирует основной текст слайда."""
        pass
    
    @abstractmethod
    def generate_bullet_points(self, context, specifications):
        """Генерирует маркированный список."""
        pass
    
    @abstractmethod
    def generate_table_content(self, context, structure, specifications):
        """Генерирует содержимое таблицы."""
        pass
    
    @abstractmethod
    def adapt_text_to_placeholder(self, text, placeholder_constraints):
        """Адаптирует текст под ограничения плейсхолдера."""
        pass
    
    @abstractmethod
    def ensure_stylistic_consistency(self, text, presentation_style):
        """Обеспечивает стилистическую согласованность текста."""
        pass
