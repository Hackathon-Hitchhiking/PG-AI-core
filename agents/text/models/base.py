from abc import ABC, abstractmethod

from agents.text.schemas import BaseTextConfig, GenerationParams


class BaseTextModel(ABC):
    @classmethod
    @abstractmethod
    def from_config(cls, config: BaseTextConfig) -> "BaseTextModel":
        """Создает экземпляр модели из конфигурации"""
        pass

    @abstractmethod
    def generate(self, prompt: str, params: GenerationParams) -> str:
        """Генерирует текст на основе промпта и параметров"""
        pass

class ModelRegistry:
    _registry: dict[type[BaseTextConfig], type[BaseTextModel]] = {}

    @classmethod
    def register(cls, config_type: type[BaseTextConfig]):
        """Декоратор для регистрации моделей в реестре"""
        def decorator(model_class: type[BaseTextModel]):
            cls._registry[config_type] = model_class
            return model_class
        return decorator

    @classmethod
    def get_model_class(cls, config: BaseTextConfig) -> type[BaseTextModel]:
        """Получает класс модели по типу конфигурации"""
        config_type = type(config)
        if config_type not in cls._registry:
            msg = f"No model registered for config type {config_type}"
            raise ValueError(msg)
        return cls._registry[config_type]
