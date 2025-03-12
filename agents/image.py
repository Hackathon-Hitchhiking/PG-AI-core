from __future__ import annotations
from abc import ABC, abstractmethod
from typing import Any, Dict, Optional, Union, Type, Literal, TypedDict
from pydantic import BaseModel, Field, ValidationError, field_validator
import logging
import importlib
from PIL import Image
import requests
import torch
from io import BytesIO
import hashlib
import json
import inspect

logger = logging.getLogger(__name__)

ModelType = Literal["hf", "diffusers", "api", "local"]

class ModelConfig(BaseModel):
    model_type: ModelType = Field(..., description="Тип модели")
    model_name: Optional[str] = Field(None, min_length=1)
    api_base: Optional[str] = Field(None, min_length=3)
    api_key: Optional[str] = Field(None, min_length=1)
    device: str = Field(default="cuda" if torch.cuda.is_available() else "cpu", min_length=1)
    torch_dtype: Literal["float16", "float32", "bfloat16"] = "float16"
    revision: Optional[str] = None
    model_kwargs: Dict[str, Any] = Field(default_factory=dict)
    pipeline_kwargs: Dict[str, Any] = Field(default_factory=dict)
    api_handler: Optional[str] = Field(None, pattern=r"^[\w\.]+\.\w+$")

    @field_validator("model_name", always=True)
    def validate_model_name(cls, v, values):
        if values.get("model_type") in ["hf", "diffusers"] and not v:
            raise ValueError("Model name is required for HF/Diffusers models")
        return v

class GenerationRequest(BaseModel):
    prompt: str = Field(..., min_length=1)
    negative_prompt: Optional[str] = None
    width: int = Field(512, ge=64, le=2048)
    height: int = Field(512, ge=64, le=2048)
    num_inference_steps: int = Field(50, ge=1, le=150)
    guidance_scale: float = Field(7.5, ge=0.0, le=20.0)
    seed: Optional[int] = Field(None, ge=0, le=2**32-1)
    output_type: Literal["pil", "latent"] = "pil"
    adapter: Optional[str] = None
    lora_weights: Optional[str] = None
    controlnet: Optional[str] = None

class ImageResult(TypedDict):
    image: Image.Image
    metadata: Dict[str, Any]

class BaseImageModel(ABC):
    @abstractmethod
    def generate(self, request: GenerationRequest) -> Image.Image:
        pass
    
    @classmethod
    @abstractmethod
    def from_config(cls, config: ModelConfig) -> BaseImageModel:
        pass

class ImageAgent:
    def __init__(self, config: Union[Dict[str, Any], ModelConfig]):
        self.config = config if isinstance(config, ModelConfig) else ModelConfig(**config)
        self.model = self._init_model()
        self.cache: Dict[str, Image.Image] = {}
        
        logger.info(f"Initialized {self.config.model_type} model: {self.config.model_name}")

    def _init_model(self) -> BaseImageModel:
        model_registry: Dict[ModelType, Type[BaseImageModel]] = {
            "hf": HuggingFaceModel,
            "diffusers": DiffusersModel,
            "api": APIModel,
            "local": LocalModel
        }
        return model_registry[self.config.model_type].from_config(self.config)

    def generate(
        self,
        request: Union[Dict[str, Any], GenerationRequest],
        cache_key: Optional[str] = None,
        fallback_strategies: list[ModelType] = ["hf", "api", "diffusers"]
    ) -> Image.Image:
        try:
            validated_request = self._validate_request(request)
            cache_key = cache_key or self._generate_cache_key(validated_request)
            
            if cached_image := self.cache.get(cache_key):
                logger.info("Using cached image")
                return cached_image
                
            image = self.model.generate(validated_request)
            self.cache[cache_key] = image
            return image
            
        except Exception as e:
            logger.error(f"Generation failed: {str(e)}")
            return self._handle_fallback(validated_request, fallback_strategies)

    def _validate_request(self, request: Union[Dict[str, Any], GenerationRequest]) -> GenerationRequest:
        if isinstance(request, GenerationRequest):
            return request
        try:
            return GenerationRequest(**request)
        except ValidationError as e:
            logger.error(f"Invalid request parameters: {str(e)}")
            raise

    def _generate_cache_key(self, request: GenerationRequest) -> str:
        return hashlib.sha256(request.json().encode()).hexdigest()

    def _handle_fallback(
        self,
        request: GenerationRequest,
        strategies: list[ModelType]
    ) -> Image.Image:
        for strategy in strategies:
            try:
                fallback_model = self._init_fallback_model(strategy)
                return fallback_model.generate(request)
            except Exception as e:
                logger.warning(f"Fallback {strategy} failed: {str(e)}")
        raise RuntimeError("All generation strategies failed")

    def _init_fallback_model(self, model_type: ModelType) -> BaseImageModel:
        return type(self)(ModelConfig(**{**self.config.dict(), "model_type": model_type})).model

class HuggingFaceModel(BaseImageModel):
    def __init__(self, pipeline: Any):
        self.pipeline = pipeline

    @classmethod
    def from_config(cls, config: ModelConfig) -> HuggingFaceModel:
        from transformers import pipeline
        return cls(
            pipeline(
                "text-to-image",
                model=config.model_name,
                device=config.device,
                torch_dtype=getattr(torch, config.torch_dtype),
                **config.pipeline_kwargs
            )
        )

    def generate(self, request: GenerationRequest) -> Image.Image:
        generator = torch.Generator(device=self.pipeline.device).manual_seed(request.seed) if request.seed else None
        result = self.pipeline(
            **request.model_dump(exclude={"seed"}),
            generator=generator
        )
        return result.images[0]

class DiffusersModel(BaseImageModel):
    def __init__(self, pipeline: Any):
        self.pipeline = pipeline

    @classmethod
    def from_config(cls, config: ModelConfig) -> DiffusersModel:
        from diffusers import DiffusionPipeline
        return cls(
            DiffusionPipeline.from_pretrained(
                config.model_name,
                torch_dtype=getattr(torch, config.torch_dtype),
                revision=config.revision,
                **config.model_kwargs
            ).to(config.device)
        )

    def generate(self, request: GenerationRequest) -> Image.Image:
        generator = torch.Generator(self.pipeline.device).manual_seed(request.seed) if request.seed else None
        return self.pipeline(**request.model_dump(exclude={"seed"}), generator=generator).images[0]

class APIModel(BaseImageModel):
    def __init__(self, config: ModelConfig):
        self.config = config

    @classmethod
    def from_config(cls, config: ModelConfig) -> APIModel:
        return cls(config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        handler = self._get_api_handler()
        return handler(request)

    def _get_api_handler(self) -> callable:
        if self.config.api_handler:
            module_path, class_name = self.config.api_handler.rsplit(".", 1)
            module = importlib.import_module(module_path)
            handler_class = getattr(module, class_name)
            if not inspect.isclass(handler_class):
                raise TypeError("API handler must be a class")
            return handler_class(self.config).handle
        return self._default_api_handler

    def _default_api_handler(self, request: GenerationRequest) -> Image.Image:
        response = requests.post(
            self.config.api_base,
            headers={"Authorization": f"Bearer {self.config.api_key}"},
            json={**request.model_dump(), **self.config.model_kwargs},
            timeout=30
        )
        response.raise_for_status()
        
        if not (image_url := response.json().get("data", [{}])[0].get("url")):
            raise ValueError("Invalid API response format")
            
        image_response = requests.get(image_url, timeout=30)
        return Image.open(BytesIO(image_response.content))

class LocalModel(BaseImageModel):
    @classmethod
    def from_config(cls, config: ModelConfig) -> LocalModel:
        raise NotImplementedError("Local models implementation required")

    def generate(self, request: GenerationRequest) -> Image.Image:
        raise NotImplementedError("Local models implementation required")