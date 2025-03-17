from pathlib import Path
from typing import Dict, Any, Optional
import yaml
from agents.text.schemas import ContentStyle, ContentTone

class PromptLibrary:
    """Библиотека шаблонов промптов для различных задач генерации."""
    
    DEFAULT_TEMPLATES = {
        'title': {
            ContentStyle.ARTICLE: "Создай информативный заголовок для статьи",
            ContentStyle.PRESENTATION: "Придумай яркий, цепляющий заголовок для презентации",
            ContentStyle.BLOG: "Напиши привлекательный заголовок для блога",
            ContentStyle.ACADEMIC: "Сформулируй академический заголовок",
            ContentStyle.MARKETING: "Создай продающий заголовок",
            ContentStyle.NEWS: "Напиши новостной заголовок"
        },
        'content': {
            ContentStyle.ARTICLE: "Напиши подробную статью",
            ContentStyle.PRESENTATION: "Создай яркий и лаконичный текст для презентации",
            ContentStyle.BLOG: "Напиши увлекательный блог-пост",
            ContentStyle.ACADEMIC: "Подготовь академический текст",
            ContentStyle.MARKETING: "Создай убедительный маркетинговый текст",
            ContentStyle.NEWS: "Напиши новостную статью"
        },
        'title_template': """
        {instruction} в {tone} стиле.
        Тема: {context}
        Требования: краткость, ясность, привлекательность.
        Заголовок должен быть на русском языке.
        """.strip(),
        'content_template': """
        {instruction} в формате {format_hint}.
        Тема: {context}
        Стиль: {tone}
        Требования:
        - Текст должен быть на русском языке
        - Структурированное изложение
        - Логичные переходы между частями
        - Соответствие выбранному стилю и тону
        - Информативность и полнота раскрытия темы
        """.strip(),
        'image_template': """
        Создай детальное описание изображения.
        Тема: {context}
        Стиль: {style}
        Тональность: {tone}
        Дополнительные требования: {requirements}
        """.strip(),
        'presentation_template': """
        Создай {section_type} для слайда презентации.
        Тема: {context}
        Стиль: {style}
        Тональность: {tone}
        Требования:
        - Лаконичность и структурированность
        - Ключевые идеи должны быть выделены
        - Соответствие общему стилю презентации
        """.strip()
    }
    
    def __init__(self, templates_path: Optional[str | Path] = None):
        """
        Инициализация библиотеки с опциональным путем к файлу шаблонов.
        
        Args:
            templates_path: Путь к YAML файлу с дополнительными шаблонами
        """
        self.templates = self.DEFAULT_TEMPLATES.copy()
        if templates_path:
            self._load_templates(templates_path)
    
    def _load_templates(self, templates_path: str | Path) -> None:
        """
        Загружает пользовательские шаблоны из YAML файла.
        
        Args:
            templates_path: Путь к YAML файлу с шаблонами
        """
        path = Path(templates_path)
        if path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                custom_templates = yaml.safe_load(f)
                self.templates.update(custom_templates)
    
    def get_prompt(self, prompt_type: str, **params: Any) -> str:
        """
        Получение форматированного промпта указанного типа.
        
        Args:
            prompt_type: Тип промпта ('title', 'content', 'image', 'presentation')
            **params: Параметры для форматирования промпта
                     style: ContentStyle - стиль контента
                     tone: ContentTone - тональность
                     context: str - контекст/тема
                     format_hint: str - подсказка формата
                     requirements: str - дополнительные требования
                     section_type: str - тип секции презентации
        
        Returns:
            str: Отформатированный промпт
        """
        style = params.get('style', ContentStyle.ARTICLE)
        tone = params.get('tone', ContentTone.PROFESSIONAL)
        
        if prompt_type == 'title':
            instruction = self.templates['title'][style]
            return self.templates['title_template'].format(
                instruction=instruction,
                tone=tone.value,
                context=params.get('context', '')
            )
        
        elif prompt_type == 'content':
            instruction = self.templates['content'][style]
            return self.templates['content_template'].format(
                instruction=instruction,
                tone=tone.value,
                context=params.get('context', ''),
                format_hint=params.get('format_hint', 'текст')
            )
        
        elif prompt_type == 'image':
            return self.templates['image_template'].format(
                context=params.get('context', ''),
                style=style.value,
                tone=tone.value,
                requirements=params.get('requirements', '')
            )
        
        elif prompt_type == 'presentation':
            return self.templates['presentation_template'].format(
                section_type=params.get('section_type', 'содержание'),
                context=params.get('context', ''),
                style=style.value,
                tone=tone.value
            )
        
        else:
            template = self.templates.get(prompt_type)
            if template:
                return template.format(**params)
            raise ValueError(f"Unknown prompt type: {prompt_type}")
    
    def add_template(self, prompt_type: str, template: str | Dict[str, str]) -> None:
        """
        Добавление нового шаблона промпта.
        
        Args:
            prompt_type: Тип промпта
            template: Шаблон или словарь шаблонов
        """
        if prompt_type in self.templates:
            if isinstance(self.templates[prompt_type], dict):
                if isinstance(template, dict):
                    self.templates[prompt_type].update(template)
                else:
                    raise ValueError(f"Template for {prompt_type} must be a dictionary")
            else:
                self.templates[prompt_type] = template
        else:
            self.templates[prompt_type] = template
