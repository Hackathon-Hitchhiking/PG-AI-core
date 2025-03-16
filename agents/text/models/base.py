from abc import ABC, abstractmethod
from typing import Type
from agents.text.schemas import GenerationParams, BaseTextConfig

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
    _registry: dict[Type[BaseTextConfig], Type[BaseTextModel]] = {}

    @classmethod
    def register(cls, config_type: Type[BaseTextConfig]):
        """Декоратор для регистрации моделей в реестре"""
        def decorator(model_class: Type[BaseTextModel]):
            cls._registry[config_type] = model_class
            return model_class
        return decorator

    @classmethod
    def get_model_class(cls, config: BaseTextConfig) -> Type[BaseTextModel]:
        """Получает класс модели по типу конфигурации"""
        config_type = type(config)
        if config_type not in cls._registry:
            raise ValueError(f"No model registered for config type {config_type}")
        return cls._registry[config_type]