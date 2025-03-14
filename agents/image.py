from __future__ import annotations

import hashlib
import importlib
import logging
from abc import ABC, abstractmethod
from typing import Any, Callable, Dict, Literal, Optional, Type, TypedDict, Union

import openai
import requests
import torch
from cachetools import LRUCache
from pydantic import BaseModel, Field, field_validator
from PIL import Image
from io import BytesIO

logger = logging.getLogger(__name__)

ModelType = Literal["diffusers", "api", "local", "openai"]


class ImageModelConfig(BaseModel):
    """
    Configuration for an image model.

    Parameters
    ----------
    model_type : ModelType
        The type of model to use ('diffusers', 'api', 'local', 'openai').
    model_name : str, optional
        Name or path of the model, required for diffusers.
    api_base : str, optional
        Base URL for API calls, if using an API-based model.
    api_key : str, optional
        API key for external services (if needed).
    device : str
        The device to run inference on ('cuda' or 'cpu').
    torch_dtype : {'float16', 'float32', 'bfloat16'}
        The torch dtype used for diffusion pipelines.
    revision : str, optional
        Model revision (if applicable).
    model_kwargs : dict
        Additional kwargs for model loading or request parameters.
    pipeline_kwargs : dict
        Additional kwargs for pipeline usage (diffusers).
    api_handler : str, optional
        Dot-path to a custom API handler class (if needed).
    """

    model_type: ModelType
    model_name: Optional[str] = Field(None, min_length=1)
    api_base: Optional[str] = Field(None, min_length=3)
    api_key: Optional[str] = Field(None, min_length=1)
    device: str = Field(default="cuda" if torch.cuda.is_available() else "cpu")
    torch_dtype: Literal["float16", "float32", "bfloat16"] = "float16"
    revision: Optional[str] = None
    model_kwargs: Dict[str, Any] = Field(default_factory=dict)
    pipeline_kwargs: Dict[str, Any] = Field(default_factory=dict)
    api_handler: Optional[str] = Field(None, pattern=r"^[\w\.]+\.\w+$")

    @field_validator("model_name")
    def validate_model_name(cls, v, values):
        if values.data.get("model_type") == "diffusers" and not v:
            raise ValueError("model_name is required when model_type='diffusers'.")
        return v


class GenerationRequest(BaseModel):
    """
    Generation request parameters.

    Parameters
    ----------
    prompt : str
        The text prompt for image generation.
    negative_prompt : str, optional
        A negative text prompt for certain pipelines.
    width : int
        Width of the generated image.
    height : int
        Height of the generated image.
    num_inference_steps : int
        Number of diffusion inference steps.
    guidance_scale : float
        Guidance scale factor for diffusion.
    seed : int, optional
        Random seed.
    output_type : {'pil', 'latent'}
        The type of output, PIL image or latent representation.
    adapter : str, optional
        Adapter specification for certain pipelines.
    lora_weights : str, optional
        LoRA weights path for certain pipelines.
    controlnet : str, optional
        ControlNet model spec for certain pipelines.
    """

    prompt: str = Field(..., min_length=1)
    negative_prompt: Optional[str] = None
    width: int = Field(512, ge=64, le=2048)
    height: int = Field(512, ge=64, le=2048)
    num_inference_steps: int = Field(50, ge=1, le=150)
    guidance_scale: float = Field(7.5, ge=0.0, le=20.0)
    seed: Optional[int] = Field(None, ge=0, le=2**32 - 1)
    output_type: Literal["pil", "latent"] = "pil"
    adapter: Optional[str] = None
    lora_weights: Optional[str] = None
    controlnet: Optional[str] = None


class ImageResult(TypedDict):
    """
    Image generation result.

    Keys
    ----
    image : PIL.Image.Image
        The generated image.
    metadata : dict
        Additional metadata (if needed).
    """

    image: Image.Image
    metadata: Dict[str, Any]


class BaseImageModel(ABC):
    """
    Abstract base class for image models.
    """

    @abstractmethod
    def generate(self, request: GenerationRequest) -> Image.Image:
        """
        Generates an image based on a GenerationRequest.

        Parameters
        ----------
        request : GenerationRequest
            Parameters for the image generation.

        Returns
        -------
        PIL.Image.Image
            Generated PIL image.
        """
        pass

    @classmethod
    @abstractmethod
    def from_config(cls, config: ImageModelConfig) -> BaseImageModel:
        """
        Constructs a model instance from the given configuration.

        Parameters
        ----------
        config : ImageModelConfig
            Model configuration object.

        Returns
        -------
        BaseImageModel
            An instance of a concrete image model class.
        """
        pass


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
            # Если задана кастомная логика (класс), грузим его
            module_path, class_name = self.config.api_handler.rsplit(".", 1)
            module = importlib.import_module(module_path)
            handler_class = getattr(module, class_name)
            handler_instance = handler_class(self.config)
            if hasattr(handler_instance, "handle") and callable(handler_instance.handle):
                return handler_instance.handle
            raise TypeError("Custom API handler must implement 'handle' method.")
        # Иначе пользуемся "умным" дефолтом
        return self._default_api_handler

    def _default_api_handler(self, request: GenerationRequest) -> Image.Image:
        method = str(self.config.model_kwargs.get("method", "POST")).upper()

        if method == "GET":
            # Пример: Pollinations (или любой сервис, куда нужно дергать GET)
            base = self.config.api_base
            if not base:
                raise ValueError("api_base is not set for GET requests.")
            full_url = f"{base}{requests.utils.quote(request.prompt)}"
            resp = requests.get(full_url, timeout=30)
            resp.raise_for_status()
            return Image.open(BytesIO(resp.content))
        else:
            # По умолчанию делаем POST и пытаемся считать JSON {"data":[{"url":"..."}]}
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
            "local": LocalModel,
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
        base = self.config.dict()
        base["model_type"] = model_type
        return ImageModelConfig(**base)


class LocalModel(BaseImageModel):
    """
    Model stub for local text-to-image generation.
    """

    @classmethod
    def from_config(cls, config: ImageModelConfig) -> LocalModel:
        raise NotImplementedError("LocalModel is not implemented.")

    def generate(self, request: GenerationRequest) -> Image.Image:
        raise NotImplementedError("LocalModel generation is not implemented.")