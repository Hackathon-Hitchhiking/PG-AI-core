from __future__ import annotations

from typing import Any, Dict, Literal, Optional, TypedDict

import torch
from pydantic import BaseModel, Field, field_validator
from PIL import Image

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
    folder_id: Optional[str] = Field(None, min_length=1)
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
    seed: Optional[int] = Field(None, ge=0, le=2 ** 32 - 1)
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
