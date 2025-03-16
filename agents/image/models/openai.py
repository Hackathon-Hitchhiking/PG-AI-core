# agents/image/models/openai.py
import base64
from typing import Any
import openai
import requests
from io import BytesIO
from PIL import Image
from schemas.image import OpenAIConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry

@ModelRegistry.register(OpenAIConfig)
class OpenAIModel(BaseImageModel):
    def __init__(self, client: Any, config: OpenAIConfig):
        self.client: openai.Client = client
        self.config: OpenAIConfig = config

    @classmethod
    def from_config(cls, config: OpenAIConfig) -> "OpenAIModel":
        client = openai.Client(api_key=config.api_key.get_secret_value())
        return cls(client, config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        response = self.client.images.generate(
            model="dall-e-3",
            prompt=request.prompt,
            size=f"{request.width}x{request.height}",
            quality=self.config.quality,
            style=self.config.style,
            response_format=self.config.response_format,
            timeout=self.config.timeout
        )

        if self.config.response_format == "url":
            return self._download_image(response.data[0].url)
        else:
            return Image.open(BytesIO(base64.b64decode(response.data[0].b64_json)))

    def _download_image(self, url: str) -> Image.Image:
        response = requests.get(url, timeout=self.config.timeout)
        response.raise_for_status()
        return Image.open(BytesIO(response.content))