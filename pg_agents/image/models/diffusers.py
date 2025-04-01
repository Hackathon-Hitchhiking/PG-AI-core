import torch

from agents.image.models.base import BaseImageModel, ModelRegistry
from agents.image.schemas import DiffusersConfig, GenerationRequest
from diffusers import DiffusionPipeline
from PIL import Image


@ModelRegistry.register(DiffusersConfig)
class DiffusersModel(BaseImageModel):
    def __init__(self, pipeline: DiffusionPipeline, config: DiffusersConfig) -> None:
        self.pipeline: DiffusionPipeline = pipeline
        self.config: DiffusersConfig = config

    @classmethod
    def from_config(cls, config: DiffusersConfig) -> 'DiffusersModel':
        pipeline = DiffusionPipeline.from_pretrained(
            config.model_name, torch_dtype=getattr(torch, config.torch_dtype), **config.pipeline_kwargs
        )

        # Применяем оптимизации
        pipeline = pipeline.to(config.device.value)

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
            generator=generator,
        ).images[0]
