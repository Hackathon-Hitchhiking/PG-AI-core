# agents/image/models/diffusers.py
from typing import Any
import torch
from PIL import Image
from schemas.image import DiffusersConfig, GenerationRequest, ModelType
from agents.image.models.base import BaseImageModel, ModelRegistry

@ModelRegistry.register(ModelType.DIFFUSERS)
class DiffusersModel(BaseImageModel):
    def __init__(self, pipeline: Any, config: DiffusersConfig):
        self.pipeline = pipeline
        self.config = config

    @classmethod
    def from_config(cls, config: DiffusersConfig) -> "DiffusersModel":
        from diffusers import AutoPipelineForText2Image  # Lazy import
        
        pipeline = AutoPipelineForText2Image.from_pretrained(
            config.model_name,
            torch_dtype=getattr(torch, config.torch_dtype),
            device=config.device,
            **config.pipeline_kwargs
        )
        return cls(pipeline, config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        generator = torch.Generator(device=self.config.device)
        if request.seed is not None:
            generator.manual_seed(request.seed)

        return self.pipeline(
            **request.model_dump(),
            generator=generator,
            **self.config.pipeline_kwargs
        ).images[0]