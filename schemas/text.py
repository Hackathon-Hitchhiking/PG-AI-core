from __future__ import annotations
from typing import Any, Dict, Optional, Literal, List, TypedDict
from pydantic import BaseModel, Field, field_validator
import torch

TextBackend = Literal["hf", "api", "yandexgpt", "langchain", "llamacpp", "llamaindex", "vllm", "transformers"]

class TextModelConfig(BaseModel):
    backend: TextBackend = Field(..., description="Тип бэкенда для текстовой модели")
    model_name: Optional[str] = Field(None, min_length=1)
    api_base: Optional[str] = Field(None, min_length=3)
    api_key: Optional[str] = Field(None, min_length=1)
    folder_id: Optional[str] = Field(None, min_length=1)
    model_path: Optional[str] = None
    device: str = Field(default="cuda" if torch.cuda.is_available() else "cpu")
    torch_dtype: Literal["auto", "float16", "float32"] = "auto"
    tokenizer_name: Optional[str] = None
    context_length: int = 4096
    temperature: float = 0.7
    max_new_tokens: int = 512
    top_p: float = 0.95
    langchain_template: Optional[str] = None
    llamacpp_params: Dict[str, Any] = Field(default_factory=dict)
    vector_store: Optional[str] = None
    quantized: bool = False
    use_safetensors: bool = True

    @field_validator("model_path")
    def validate_model_path(cls, v, values):
        backend = values.data.get("backend")
        if backend in ["llamacpp", "transformers"] and not v:
            raise ValueError("Model path is required for this backend")
        return v

class GenerationParams(TypedDict):
    temperature: float
    max_new_tokens: int
    top_p: float
    repetition_penalty: float
    stop_sequences: List[str]