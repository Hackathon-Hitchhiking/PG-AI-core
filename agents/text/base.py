from abc import ABC, abstractmethod

from schemas.text import GenerationParams, TextModelConfig

class BaseTextModel(ABC):
    @abstractmethod
    def generate(self, prompt: str, params: GenerationParams) -> str:
        pass

    @classmethod
    @abstractmethod
    def from_config(cls, config: TextModelConfig) -> "BaseTextModel":
        pass