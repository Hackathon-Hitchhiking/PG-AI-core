from __future__ import annotations
from enum import Enum
from typing import Any, TypedDict, Literal
from pydantic import BaseModel, Field, ConfigDict, model_validator
import torch

class DeviceType(str, Enum):
    AUTO = "auto"
    CUDA = "cuda"
    CPU = "cpu"
    MPS = "mps"
    TPU = "tpu"

class BaseTextConfig(BaseModel):
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

class HFConfig(BaseTextConfig):
    """Конфигурация для Hugging Face моделей"""
    model_name: str = Field(
        ...,
        min_length=3,
        examples=["gpt-3.5-turbo"]
    )
    torch_dtype: Literal["auto", "float16", "float32"] = "auto"
    revision: str | None = None
    tokenizer_name: str | None = None
    context_length: int = 4096
    pipeline_kwargs: dict[str, Any] = Field(
        default_factory=dict,
        description="Дополнительные параметры для пайплайна"
    )

    @model_validator(mode="after")
    def validate_device(self) -> "HFConfig":
        if self.device == DeviceType.AUTO:
            if torch.cuda.is_available():
                self.device = DeviceType.CUDA
            elif torch.backends.mps.is_available():
                self.device = DeviceType.MPS
            else:
                self.device = DeviceType.CPU
        return self

class APIConfig(BaseTextConfig):
    """Конфигурация для произвольного API"""
    api_base: str = Field(
        ...,
        min_length=8,
        examples=["https://api.example.com/v1/generate"]
    )
    api_key: str = Field(
        ...,
        min_length=32,
        description="API ключ для доступа к API"
    )
    headers: dict[str, str] = Field(
        default_factory=dict,
        description="Кастомные заголовки HTTP-запроса"
    )
    model_name: str | None = None
    response_format: Literal["text", "json"] = "text"
    response_path: list[str | int] = ["choices", 0, "text"]

class LangChainConfig(BaseTextConfig):
    """Конфигурация для LangChain моделей"""
    model_name: str = Field(
        ...,
        min_length=3,
        examples=["gpt-3.5-turbo"]
    )
    langchain_template: str = Field(
        default="{input}",
        min_length=1,
        examples=["{input}"]
    )
    context_length: int = 4096
    chain_type: Literal["llm", "stuff", "map_reduce"] = "llm"
    memory_type: Literal["buffer", "summary", "conversation"] | None = None

class LlamaCppConfig(BaseTextConfig):
    """Конфигурация для LlamaCpp моделей"""
    model_path: str = Field(
        ...,
        min_length=1,
        examples=["/path/to/llama/model"]
    )
    context_length: int = 4096
    n_gpu_layers: int = -1
    main_gpu: int = 0
    tensor_split: list[float] | None = None
    llamacpp_params: dict[str, Any] = Field(
        default_factory=dict,
        description="Дополнительные параметры для LlamaCpp"
    )

class LlamaIndexConfig(BaseTextConfig):
    """Конфигурация для LlamaIndex моделей"""
    model_name: str = Field(
        ...,
        min_length=3,
        examples=["gpt-3.5-turbo"]
    )
    vector_store: str = Field(
        ...,
        min_length=1,
        examples=["/path/to/vector/store"]
    )
    similarity_top_k: int = 3
    response_mode: Literal["compact", "tree", "refine"] = "compact"
    chunk_size: int = 1024
    chunk_overlap: int = 20

class vLLMConfig(BaseTextConfig):
    """Конфигурация для vLLM моделей"""
    model_path: str = Field(
        ...,
        min_length=1,
        examples=["/path/to/vllm/model"]
    )
    quantized: bool = False
    tensor_parallel_size: int = 1
    gpu_memory_utilization: float = 0.9
    context_length: int = 4096
    max_model_len: int | None = None
    enforce_eager: bool = False

class TransformersConfig(BaseTextConfig):
    """Конфигурация для Transformers моделей"""
    model_path: str = Field(
        ...,
        min_length=1,
        examples=["/path/to/transformers/model"]
    )
    use_safetensors: bool = True
    tokenizer_name: str | None = None
    context_length: int = 4096
    load_in_8bit: bool = False
    trust_remote_code: bool = False
    use_auth_token: bool = False

class YandexConfig(BaseTextConfig):
    """Конфигурация для Yandex GPT"""
    folder_id: str = Field(
        ...,
        min_length=1,
        examples=["b1gvi7tuub11s7aeu0ul"]
    )
    api_key: str = Field(
        ...,
        min_length=32,
        description="API ключ для доступа к YandexGPT"
    )
    context_length: int = 4096
    model_uri: str = Field(
        default="yandexgpt",
        examples=["yandexgpt-lite", "yandexgpt"]
    )

class ContentStyle(str, Enum):
    ARTICLE = "article"      # Длинная, детальная статья
    PRESENTATION = "presentation"  # Краткая, яркая подача
    BLOG = "blog"           # Неформальный, средний объем
    ACADEMIC = "academic"   # Научный стиль
    MARKETING = "marketing" # Продающий текст
    NEWS = "news"          # Новостной формат

class ContentTone(str, Enum):
    FORMAL = "formal"
    CASUAL = "casual"
    PROFESSIONAL = "professional"
    FRIENDLY = "friendly"
    ENTHUSIASTIC = "enthusiastic"

class GenerationParams(TypedDict, total=False):
    prompt: str
    temperature: float = 0.7
    max_new_tokens: int = 512
    min_new_tokens: int | None = None
    top_p: float = 0.95
    top_k: int = 50
    repetition_penalty: float = 1.1
    stop_sequences: list[str] = ["\n\n"]
    do_sample: bool = True
    num_beams: int = 1
    presence_penalty: float = 0.0
    frequency_penalty: float = 0.0
    truncate: bool = False
    skip_special_tokens: bool = True
    style: ContentStyle = ContentStyle.ARTICLE
    tone: ContentTone = ContentTone.PROFESSIONAL
    length: Literal["short", "medium", "long"] = "medium"