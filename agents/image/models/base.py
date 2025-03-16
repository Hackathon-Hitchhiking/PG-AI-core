from abc import ABC, abstractmethod
from typing import Type, Any
from pydantic import BaseModel
from PIL import Image

class BaseImageModel(ABC):
    @classmethod
    @abstractmethod
    def from_config(cls, config: BaseModel) -> "BaseImageModel":
        pass

    @abstractmethod
    def generate(self, request: Any) -> Image.Image:
        pass

class ModelRegistry:
    _registry: dict[Type[BaseModel], Type[BaseImageModel]] = {}

    @classmethod
    def register(cls, config_type: Type[BaseModel]):
        def decorator(model_class: Type[BaseImageModel]):
            cls._registry[config_type] = model_class
            return model_class
        return decorator

    @classmethod
    def get_model_class(cls, config: BaseModel) -> Type[BaseImageModel]:
        config_type = type(config)
        if config_type not in cls._registry:
            raise ValueError(f"No model registered for config type {config_type}")
        return cls._registry[config_type]