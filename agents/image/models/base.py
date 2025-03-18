from abc import ABC, abstractmethod
from typing import Any

from PIL import Image
from pydantic import BaseModel


class BaseImageModel(ABC):
    @classmethod
    @abstractmethod
    def from_config(cls, config: BaseModel) -> "BaseImageModel":
        pass

    @abstractmethod
    def generate(self, request: Any) -> Image.Image:
        pass

class ModelRegistry:
    _registry: dict[type[BaseModel], type[BaseImageModel]] = {}

    @classmethod
    def register(cls, config_type: type[BaseModel]):
        def decorator(model_class: type[BaseImageModel]):
            cls._registry[config_type] = model_class
            return model_class
        return decorator

    @classmethod
    def get_model_class(cls, config: BaseModel) -> type[BaseImageModel]:
        config_type = type(config)
        if config_type not in cls._registry:
            msg = f"No model registered for config type {config_type}"
            raise ValueError(msg)
        return cls._registry[config_type]
