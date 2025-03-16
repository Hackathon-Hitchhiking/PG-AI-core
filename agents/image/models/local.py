# agents/image/models/local.py
import torch
from PIL import Image
from typing import Any
from schemas.image import LocalModelConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry, ModelType

@ModelRegistry.register(ModelType.LOCAL)
class LocalModel(BaseImageModel):
    def __init__(self, config: LocalModelConfig):
        self.config = config
        self.model = self._load_model()
        self.device = torch.device(config.device)

    @classmethod
    def from_config(cls, config: LocalModelConfig) -> "LocalModel":
        return cls(config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        generator = torch.Generator(device=self.device)
        if request.seed is not None:
            generator.manual_seed(request.seed)

        with torch.inference_mode():
            result = self.model(
                prompt=request.prompt,
                negative_prompt=request.negative_prompt,
                width=request.width,
                height=request.height,
                num_inference_steps=request.num_inference_steps,
                guidance_scale=request.guidance_scale,
                generator=generator
            )

        return result.images[0]

    def _load_model(self) -> Any:
        from diffusers import DiffusionPipeline

        return DiffusionPipeline.from_pretrained(
            self.config.model_path,
            torch_dtype=torch.float16 if self.config.use_fp16 else torch.float32,
            cache_dir=self.config.local_cache_dir,
            device_map="auto" if self.config.device == "auto" else None
        ).to(self.device)