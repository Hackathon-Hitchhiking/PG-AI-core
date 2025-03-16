import torch
from PIL import Image
from typing import Any
from schemas.image import DiffusersConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry
from diffusers import AutoPipelineForText2Image

@ModelRegistry.register(DiffusersConfig)
class DiffusersModel(BaseImageModel):
    def __init__(self, pipeline: Any, config: DiffusersConfig):
        self.pipeline: AutoPipelineForText2Image = pipeline
        self.config: DiffusersConfig = config

    @classmethod
    def from_config(cls, config: DiffusersConfig) -> "DiffusersModel":
        from diffusers import AutoPipelineForText2Image  # Lazy import
        
        torch_dtype = getattr(torch, config.torch_dtype)
        device = config.device.value  # Получаем строковое значение enum

        pipeline = AutoPipelineForText2Image.from_pretrained(
            config.model_name,
            torch_dtype=torch_dtype,
            revision=config.revision,
            **config.pipeline_kwargs
        ).to(device)

        if config.enable_xformers:
            pipeline.enable_xformers_memory_efficient_attention()

        return cls(pipeline, config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        generator = torch.Generator(device=self.config.device.value)
        if request.seed is not None:
            generator.manual_seed(request.seed)

        return self.pipeline(
            prompt=request.prompt,
            negative_prompt=request.negative_prompt,
            width=request.width,
            height=request.height,
            num_inference_steps=request.num_inference_steps,
            guidance_scale=request.guidance_scale,
            generator=generator
        ).images[0]