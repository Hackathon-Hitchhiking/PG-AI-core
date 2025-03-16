from __future__ import annotations

import hashlib
import importlib
import logging
from typing import Any, Callable, Dict, Optional, Type, Union

import openai
import requests
import torch
from cachetools import LRUCache
from PIL import Image
from io import BytesIO
from yandex_cloud_ml_sdk import YCloudML
from .base import BaseImageModel
from schemas.image import GenerationRequest, ImageModelConfig, ModelType

logger = logging.getLogger(__name__)


class DiffusersModel(BaseImageModel):
    """
    Model for text-to-image generation using Diffusers.
    """

    def __init__(self, pipeline: Any, device: str):
        self.pipeline = pipeline
        self.device = device

    @classmethod
    def from_config(cls, config: ImageModelConfig) -> DiffusersModel:
        from diffusers import AutoPipelineForText2Image

        pipeline = AutoPipelineForText2Image.from_pretrained(
            config.model_name,
            device=config.device,
            torch_dtype=getattr(torch, config.torch_dtype),
            revision=config.revision,
            **config.pipeline_kwargs
        )
        return cls(pipeline, device=config.device)

    def generate(self, request: GenerationRequest) -> Image.Image:
        generator = None
        if request.seed is not None:
            generator = torch.Generator(device=self.device).manual_seed(request.seed)
        result = self.pipeline(
            prompt=request.prompt,
            negative_prompt=request.negative_prompt,
            width=request.width,
            height=request.height,
            num_inference_steps=request.num_inference_steps,
            guidance_scale=request.guidance_scale,
            generator=generator,
            **self._additional_args(request)
        )
        return result.images[0]

    def _additional_args(self, request: GenerationRequest) -> Dict[str, Any]:
        """
        Prepares additional arguments for the pipeline call.

        Returns
        -------
        dict
            Extra parameters that can be passed to the pipeline.
        """
        extra = {}
        if request.adapter is not None:
            extra["adapter"] = request.adapter
        if request.lora_weights is not None:
            extra["lora_weights"] = request.lora_weights
        if request.controlnet is not None:
            extra["controlnet"] = request.controlnet
        return extra


class APIModel(BaseImageModel):
    """
    Model for text-to-image generation through an external API.
    Supports GET or POST via 'method' in model_kwargs.
    """

    def __init__(self, config: ImageModelConfig):
        self.config = config
        self.api_handler: Callable[[GenerationRequest], Image.Image] = self._get_api_handler()

    @classmethod
    def from_config(cls, config: ImageModelConfig) -> APIModel:
        return cls(config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        return self.api_handler(request)

    def _get_api_handler(self) -> Callable[[GenerationRequest], Image.Image]:
        if self.config.api_handler:
            module_path, class_name = self.config.api_handler.rsplit(".", 1)
            module = importlib.import_module(module_path)
            handler_class = getattr(module, class_name)
            handler_instance = handler_class(self.config)
            if hasattr(handler_instance, "handle") and callable(handler_instance.handle):
                return handler_instance.handle
            raise TypeError("Custom API handler must implement 'handle' method.")
        return self._default_api_handler

    def _default_api_handler(self, request: GenerationRequest) -> Image.Image:
        method = str(self.config.model_kwargs.get("method", "POST")).upper()

        if method == "GET":
            base = self.config.api_base
            if not base:
                raise ValueError("api_base is not set for GET requests.")
            full_url = f"{base}{requests.utils.quote(request.prompt)}"
            resp = requests.get(full_url, timeout=30)
            resp.raise_for_status()
            return Image.open(BytesIO(resp.content))
        else:
            if not self.config.api_base or not self.config.api_key:
                raise ValueError("api_base/api_key not set for POST requests.")

            body = {**request.dict(), **self.config.model_kwargs}
            headers = {"Authorization": f"Bearer {self.config.api_key}"}
            resp = requests.post(self.config.api_base, json=body, headers=headers, timeout=30)
            resp.raise_for_status()

            data = resp.json()
            image_url = data.get("data", [{}])[0].get("url")
            if not image_url:
                raise ValueError("Missing 'url' in JSON response.")
            img_resp = requests.get(image_url, timeout=30)
            img_resp.raise_for_status()
            return Image.open(BytesIO(img_resp.content))


class OpenAIModel(BaseImageModel):
    """
    Model for OpenAI's DALL-E (or similar) text-to-image generation.
    """

    def __init__(self, config: ImageModelConfig):
        self.config = config

    @classmethod
    def from_config(cls, config: ImageModelConfig) -> OpenAIModel:
        openai.api_key = config.api_key
        return cls(config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        sz = f"{request.width}x{request.height}"
        resp = openai.Image.create(
            prompt=request.prompt,
            n=1,
            size=sz
        )
        url = resp["data"][0]["url"]
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        return Image.open(BytesIO(r.content))


class YandexAIModel(BaseImageModel):
    def __init__(self, config: ImageModelConfig, client):
        self.config = config
        self.client = client

    @classmethod
    def from_config(cls, config: ImageModelConfig) -> YandexAIModel:
        sdk = YCloudML(folder_id=config.folder_id, auth=config.api_key)
        client = sdk.models.yandex_art()
        return cls(config, client)

    def generate(self, request: GenerationRequest) -> Image.Image:
        result = self.client.run(request.prompt)

        return result[0]

class ModelManager:
    """
    Manages loaded models to avoid repeated initializations.
    """

    def __init__(self):
        self._models: Dict[Any, BaseImageModel] = {}

    def get_or_create(self, config: ImageModelConfig) -> BaseImageModel:
        key = (
            config.model_type,
            config.model_name,
            config.api_base,
            config.api_key,
            config.device,
            config.revision,
            frozenset(config.model_kwargs.items()),
            frozenset(config.pipeline_kwargs.items()),
        )
        if key not in self._models:
            self._models[key] = self._create_model(config)
        return self._models[key]

    def _create_model(self, config: ImageModelConfig) -> BaseImageModel:
        registry: Dict[ModelType, Type[BaseImageModel]] = {
            "diffusers": DiffusersModel,
            "api": APIModel,
            "openai": OpenAIModel,
        }
        return registry[config.model_type].from_config(config)


class ImageAgent:
    """
    Main class that handles text-to-image generation with caching and fallback.
    """

    def __init__(
        self,
        config: Union[ImageModelConfig, Dict[str, Any]],
        fallback_strategies: list[ModelType] = None,
        cache_maxsize: int = 100
    ):
        """
        Parameters
        ----------
        config : ImageModelConfig or dict
            Model configuration.
        fallback_strategies : list of ModelType, optional
            Ordered list of fallback model types.
        cache_maxsize : int
            Maximum number of cached items (LRU).
        """
        if isinstance(config, dict):
            self.config = ImageModelConfig(**config)
        else:
            self.config = config
        self.fallback_strategies = fallback_strategies or ["diffusers", "api", "openai", "local"]
        self.manager = ModelManager()
        self.main_model = self.manager.get_or_create(self.config)
        self.cache = LRUCache(maxsize=cache_maxsize)
        logger.info(f"Initialized ImageAgent with main model_type={self.config.model_type}")

    def generate(
        self,
        request: Union[GenerationRequest, Dict[str, Any]],
        cache_key: Optional[str] = None
    ) -> Image.Image:
        """
        Generates an image.

        Parameters
        ----------
        request : GenerationRequest or dict
            A request model or dictionary of generation parameters.
        cache_key : str, optional
            Key used to store/retrieve the image from cache.

        Returns
        -------
        PIL.Image.Image
            Generated image.
        """
        req = self._validate_request(request)
        ck = cache_key or self._cache_key(req)
        if ck in self.cache:
            return self.cache[ck]
        try:
            img = self.main_model.generate(req)
            self.cache[ck] = img
            return img
        except Exception as e:
            logger.error(f"Main model failed: {e}")
            return self._fallback(req, ck)

    def _validate_request(self, request: Union[GenerationRequest, Dict[str, Any]]) -> GenerationRequest:
        """
        Validates the request using pydantic.

        Parameters
        ----------
        request : GenerationRequest or dict
            The generation parameters.

        Returns
        -------
        GenerationRequest
            Validated request.

        Raises
        ------
        ValidationError
            If validation fails.
        """
        if isinstance(request, GenerationRequest):
            return request
        return GenerationRequest(**request)

    def _cache_key(self, request: GenerationRequest) -> str:
        """
        Creates a cache key.

        Parameters
        ----------
        request : GenerationRequest
            The request for which to create a key.

        Returns
        -------
        str
            Cache key.
        """
        return hashlib.sha256(request.model_dump_json().encode()).hexdigest()

    def _fallback(self, request: GenerationRequest, cache_key: str) -> Image.Image:
        """
        Performs fallback generation.

        Parameters
        ----------
        request : GenerationRequest
            Generation request.
        cache_key : str
            Cache key for storing results.

        Returns
        -------
        PIL.Image.Image
            Generated image.

        Raises
        ------
        RuntimeError
            If all fallback strategies fail.
        """
        original_type = self.config.model_type
        attempt_order = [t for t in self.fallback_strategies if t != original_type]
        for t in attempt_order:
            try:
                new_cfg = self._copy_config_with_type(t)
                fallback_model = self.manager.get_or_create(new_cfg)
                img = fallback_model.generate(request)
                self.cache[cache_key] = img
                return img
            except Exception as e:
                logger.warning(f"Fallback '{t}' failed: {e}")
        raise RuntimeError("All generation strategies have failed.")

    def _copy_config_with_type(self, model_type: ModelType) -> ImageModelConfig:
        """
        Copies current config and changes the model_type.

        Parameters
        ----------
        model_type : ModelType
            Desired fallback model type.

        Returns
        -------
        ImageModelConfig
            New config object with the updated type.
        """
        base = self.config.model_dump()
        base["model_type"] = model_type
        return ImageModelConfig(**base)