from __future__ import annotations
from enum import Enum
from typing import Any, Literal
from pydantic import (
    BaseModel, 
    Field, 
    model_validator,
    ValidationError,
    ConfigDict,
    SecretStr
)
import torch

class DeviceType(str, Enum):
    AUTO = "auto"
    CUDA = "cuda"
    CPU = "cpu"
    MPS = "mps"
    TPU = "tpu"

class BaseImageConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    device: DeviceType = Field(
        default=DeviceType.AUTO,
        description="Аппаратная платформа для выполнения вычислений"
    )
    timeout: int = Field(default=30, ge=10, le=120)

    def _detect_device(self) -> DeviceType:
        if torch.cuda.is_available():
            return DeviceType.CUDA
        elif torch.backends.mps.is_available():
            return DeviceType.MPS
        return DeviceType.CPU

class DiffusersConfig(BaseImageConfig):
    """Конфигурация для локальных моделей через Diffusers"""
    model_name: str = Field(
        ...,
        min_length=3,
        examples=["stabilityai/stable-diffusion-xl-base-1.0"]
    )
    torch_dtype: Literal["float16", "float32", "bfloat16"] = "float16"
    revision: str | None = None
    enable_xformers: bool = False
    pipeline_kwargs: dict[str, Any] = Field(
        default_factory=dict,
        description="Дополнительные параметры для диффузионного пайплайна"
    )

    @model_validator(mode="after")
    def validate_device(self) -> "DiffusersConfig":
        if self.device == DeviceType.AUTO:
            if torch.cuda.is_available():
                self.device = DeviceType.CUDA
            elif torch.backends.mps.is_available():
                self.device = DeviceType.MPS
            else:
                self.device = DeviceType.CPU
        return self

class YandexConfig(BaseImageConfig):
    """Конфигурация для Yandex Cloud Image Generation API"""
    folder_id: str = Field(
        ...,
        min_length=1,
        examples=["b1gvi7tuub11s7aeu0ul"]
    )
    iam_token: SecretStr = Field(
        ...,
        min_length=32,
        description="IAM токен для аутентификации в Yandex Cloud"
    )
    model_uri: str | None = Field(
        default=None,
        examples=["ajs://samodelkin/prod/art/1"]
    )
    use_preview: bool = Field(
        default=True,
        description="Использовать preview-версию API"
    )

class OpenAIConfig(BaseImageConfig):
    """Конфигурация для OpenAI DALL-E"""
    api_key: SecretStr = Field(
        ...,
        min_length=32,
        description="API ключ для доступа к OpenAI"
    )
    quality: Literal["standard", "hd"] = "standard"
    style: Literal["vivid", "natural"] | None = None
    response_format: Literal["url", "b64_json"] = "url"

class APIConfig(BaseImageConfig):
    """Конфигурация для произвольного API"""
    api_base: str = Field(
        ...,
        min_length=8,
        examples=["https://api.example.com/v1/generate"]
    )
    api_key: SecretStr | None = None
    method: Literal["POST", "GET"] = "POST"
    headers: dict[str, str] = Field(
        default_factory=dict,
        description="Кастомные заголовки HTTP-запроса"
    )
    response_json_path: list[str | int] = ["data", 0, "url"]

class GenerationRequest(BaseModel):
    """Универсальный запрос для генерации изображений"""
    prompt: str = Field(
        ..., 
        min_length=1,
        max_length=2000,
        examples=["A futuristic cityscape at sunset"]
    )
    negative_prompt: str | None = Field(
        default=None,
        examples=["blurry, low quality, text"]
    )
    width: int = Field(
        default=1024,
        ge=256,
        le=2048
    )
    height: int = Field(
        default=1024,
        ge=256,
        le=2048
    )
    num_inference_steps: int = Field(
        default=30,
        ge=10,
        le=150
    )
    guidance_scale: float = Field(
        default=7.5,
        ge=0.0,
        le=20.0
    )
    seed: int | None = Field(
        default=None,
        ge=0,
        le=2**63-1
    )

    model_config = ConfigDict(
        json_schema_extra={
            "example": {
                "prompt": "A cyberpunk city at night",
                "width": 1024,
                "height": 768,
                "num_inference_steps": 40
            }
        }
    )