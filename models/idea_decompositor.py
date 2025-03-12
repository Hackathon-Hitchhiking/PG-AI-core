class IdeaDecompositor:
    """
    Центральный агент для декомпозиции пользовательского запроса на структурированные задачи.
    Координирует работу TextGen и ImageGen, управляет контекстом презентации.
    """
    
    def __init__(self, model_provider, context_manager, config):
        """Инициализирует агента с провайдером моделей и менеджером контекста."""
        pass
    
    def analyze_request(self, user_prompt, document_content=None):
        """Анализирует пользовательский запрос и извлекает основные идеи."""
        pass
    
    def create_presentation_plan(self, analyzed_request, template_structure):
        """Создает структурированный план презентации на основе анализа запроса."""
        pass
    
    def generate_slide_specifications(self, presentation_plan, template_structure):
        """Генерирует детальные спецификации для каждого слайда."""
        pass
    
    def coordinate_content_generation(self, slide_specs, context):
        """Координирует генерацию текста и изображений для слайдов."""
        pass
    
    def validate_generated_content(self, generated_content, slide_specs):
        """Проверяет соответствие сгенерированного контента спецификациям."""
        pass
    
    def refine_content(self, content, feedback):
        """Улучшает сгенерированный контент на основе обратной связи."""
        pass
    
    def maintain_presentation_coherence(self, slides):
        """Обеспечивает связность и последовательность слайдов презентации."""
        pass
