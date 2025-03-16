from __future__ import annotations
from typing import Annotated, Any, Literal, Union
from enum import Enum
from pydantic import (
    BaseModel, 
    Field, 
    model_validator,
    ConfigDict,
    ValidationError
)
import torch
import sys

class DeviceType(str, Enum):
    CUDA = "cuda"
    CPU = "cpu"
    MPS = "mps"  # Apple Silicon
    TPU = "tpu"  # Google TPU
    AUTO = "auto"

class ModelType(str, Enum):
    DIFFUSERS = "diffusers"
    API = "api"
    OPENAI = "openai"
    YANDEX = "yandex"
    LOCAL = "local"

class BaseImageConfig(BaseModel):
    model_type: ModelType
    model_name: str | None = Field(
        default=None, 
        min_length=1,
        examples=["stabilityai/stable-diffusion-2-1"]
    )
    device: DeviceType = Field(
        default=DeviceType.AUTO,
        description="Hardware device for inference",
        examples=["cuda", "mps"]
    )
    timeout: int = Field(
        default=30, 
        ge=10, 
        le=120,
        description="Inference timeout in seconds"
    )
    
    @model_validator(mode="after")
    def validate_device(self) -> "BaseImageConfig":
        if self.device == DeviceType.AUTO:
            if torch.cuda.is_available():
                self.device = DeviceType.CUDA
            elif torch.backends.mps.is_available():
                self.device = DeviceType.MPS
            else:
                self.device = DeviceType.CPU
        return self

class DiffusersConfig(BaseImageConfig):
    model_type: Literal[ModelType.DIFFUSERS] = ModelType.DIFFUSERS
    torch_dtype: Literal["float16", "float32", "bfloat16"] = Field(
        default="float16",
        description="Precision for tensors"
    )
    revision: str | None = Field(
        default=None,
        examples=["fp16", "main"]
    )
    enable_xformers: bool = Field(
        default=False,
        description="Enable memory-efficient attention"
    )
    pipeline_kwargs: dict[str, Any] = Field(
        default_factory=dict,
        examples=[{"safety_checker": None}]
    )

    @model_validator(mode="after")
    def validate_diffusers(self):
        # CUDA validation
        if self.device == DeviceType.CUDA:
            if not torch.cuda.is_available():
                raise ValueError("CUDA device requested but not available")
            if self.torch_dtype == "float16" and torch.cuda.get_device_capability()[0] < 7:
                raise ValueError("float16 requires compute capability >= 7.0")
        
        # MPS validation
        if self.device == DeviceType.MPS:
            if not torch.backends.mps.is_available():
                raise ValueError("MPS requested but not available")
            if sys.platform != "darwin":
                raise ValueError("MPS is only available on macOS")
        
        # TPU validation
        if self.device == DeviceType.TPU:
            try:
                import torch_xla
            except ImportError:
                raise ValueError("TPU requires torch_xla package")
        
        return self

class OpenAIConfig(BaseImageConfig):
    model_type: Literal[ModelType.OPENAI] = ModelType.OPENAI
    api_version: str = Field(
        default="v1",
        pattern=r"^v\d+$",
        examples=["v1", "v2"]
    )
    quality: Literal["standard", "hd"] = Field(
        default="standard",
        description="Image quality tier"
    )
    style: Literal["vivid", "natural"] | None = Field(
        default=None,
        description="Image style preference"
    )
    response_format: Literal["url", "b64_json"] = Field(
        default="url",
        description="Response format"
    )
    
    @model_validator(mode="after")
    def validate_openai_device(self):
        if self.device not in [DeviceType.CPU, DeviceType.AUTO]:
            raise ValidationError("OpenAI models don't support device selection")
        return self

class APIConfig(BaseImageConfig):
    model_type: Literal[ModelType.API] = ModelType.API
    api_base: str = Field(
        ...,
        min_length=8,
        examples=["https://api.example.com/v1"]
    )
    api_key: str | None = Field(
        default=None,
        min_length=16,
        description="API authentication key"
    )
    method: Literal["POST", "GET"] = Field(
        default="POST",
        description="HTTP method"
    )
    max_retries: int = Field(
        default=3,
        ge=0,
        le=10
    )
    response_json_path: list[str] = Field(
        default=["data", "0", "url"],
        description="JSON path to image URL"
    )
    
    @model_validator(mode="after")
    def validate_api_device(self):
        if self.device != DeviceType.AUTO:
            raise ValidationError("API models don't support device selection")
        return self

class YandexConfig(BaseImageConfig):
    model_type: Literal[ModelType.YANDEX] = ModelType.YANDEX
    folder_id: str = Field(
        ...,
        min_length=1,
        examples=["b1gexample123"]
    )
    iam_token: str = Field(
        ...,
        min_length=32,
        description="Yandex Cloud IAM token"
    )
    model_uri: str = Field(
        examples=["ajs://example/model/1"]
    )
    use_preview: bool = Field(
        default=True,
        description="Use preview version of API"
    )
    
    @model_validator(mode="after")
    def validate_yandex_device(self):
        if self.device != DeviceType.AUTO:
            raise ValidationError("Yandex models don't support device selection")
        return self

class LocalModelConfig(BaseImageConfig):
    model_type: Literal[ModelType.LOCAL] = ModelType.LOCAL
    model_path: str = Field(
        ...,
        examples=["/models/stable_diffusion"]
    )
    use_fp16: bool = Field(
        default=True,
        description="Use half-precision"
    )
    compile_model: bool = Field(
        default=False,
        description="Use torch.compile() optimization"
    )
    local_cache_dir: str | None = Field(
        default=None,
        description="Cache directory for models"
    )
    
    @model_validator(mode="after")
    def validate_local_device(self):
        if self.device == DeviceType.TPU:
            raise ValidationError("TPU not supported for local models")
        return self

ImageModelConfig = Annotated[
    Union[
        DiffusersConfig,
        OpenAIConfig,
        APIConfig,
        YandexConfig,
        LocalModelConfig
    ],
    Field(discriminator="model_type")
]

class GenerationRequest(BaseModel):
    prompt: str = Field(
        ..., 
        min_length=1,
        max_length=2000,
        examples=["A beautiful sunset over mountains"]
    )
    negative_prompt: str | None = Field(
        default=None,
        examples=["blurry, low quality"]
    )
    width: int = Field(
        default=1024,
        ge=256,
        le=2048,
        description="Image width in pixels"
    )
    height: int = Field(
        default=1024,
        ge=256,
        le=2048,
        description="Image height in pixels"
    )
    num_inference_steps: int = Field(
        default=30,
        ge=10,
        le=150,
        description="Number of diffusion steps"
    )
    guidance_scale: float = Field(
        default=7.5,
        ge=0.0,
        le=20.0,
        description="Classifier-free guidance scale"
    )
    seed: int | None = Field(
        default=None,
        ge=0,
        le=2**63-1,
        examples=[123456789]
    )
    
    model_config = ConfigDict(
        extra="forbid",
        json_schema_extra={
            "example": {
                "prompt": "A cyberpunk cityscape at night",
                "width": 1024,
                "height": 768,
                "num_inference_steps": 40,
                "guidance_scale": 8.5
            }
        }
    )