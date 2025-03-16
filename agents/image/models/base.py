# agents/image/models/base.py
from abc import ABC, abstractmethod
from PIL import Image
from schemas.image import GenerationRequest, ImageModelConfig, ModelType

class BaseImageModel(ABC):
    @abstractmethod
    def generate(self, request: GenerationRequest) -> Image.Image:
        pass

    @classmethod
    @abstractmethod
    def from_config(cls, config: ImageModelConfig) -> "BaseImageModel":
        pass

class ModelRegistry:
    _registry: dict[ModelType, type[BaseImageModel]] = {}

    @classmethod
    def register(cls, model_type: ModelType):
        def decorator(model_cls: type[BaseImageModel]):
            cls._registry[model_type] = model_cls
            return model_cls
        return decorator

    @classmethod
    def get_model_class(cls, model_type: ModelType) -> type[BaseImageModel]:
        if model_type not in cls._registry:
            raise ValueError(f"Model type {model_type} not registered")
        return cls._registry[model_type]