# agents/image/models/openai.py
import base64
import requests
from PIL import Image
from io import BytesIO
import openai
from schemas.image import OpenAIConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry, ModelType

@ModelRegistry.register(ModelType.OPENAI)
class OpenAIModel(BaseImageModel):
    def __init__(self, config: OpenAIConfig):
        self.config = config
        self.client = openai.Client(api_key=config.api_key.get_secret_value())

    @classmethod
    def from_config(cls, config: OpenAIConfig) -> "OpenAIModel":
        return cls(config)

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
            image_url = response.data[0].url
            return self._download_image(image_url)
        else:
            return Image.open(BytesIO(base64.b64decode(response.data[0].b64_json)))

    def _download_image(self, url: str) -> Image.Image:
        response = requests.get(url, timeout=self.config.timeout)
        response.raise_for_status()
        return Image.open(BytesIO(response.content))