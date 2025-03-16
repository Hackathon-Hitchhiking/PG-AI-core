# agents/image/models/api.py
import requests
from PIL import Image
from io import BytesIO
from pydantic import HttpUrl
from schemas.image import APIConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry, ModelType

@ModelRegistry.register(ModelType.API)
class APIModel(BaseImageModel):
    def __init__(self, config: APIConfig):
        self.config = config
        self.session = requests.Session()
        adapter = requests.adapters.HTTPAdapter(max_retries=config.max_retries)
        self.session.mount("https://", adapter)

    @classmethod
    def from_config(cls, config: APIConfig) -> "APIModel":
        return cls(config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        if self.config.method == "GET":
            return self._handle_get(request)
        return self._handle_post(request)

    def _handle_get(self, request: GenerationRequest) -> Image.Image:
        url = HttpUrl(f"{self.config.api_base}?prompt={requests.utils.quote(request.prompt)}")
        response = self.session.get(
            url,
            headers={"Authorization": f"Bearer {self.config.api_key}"} if self.config.api_key else None,
            timeout=self.config.timeout
        )
        return self._process_response(response)

    def _handle_post(self, request: GenerationRequest) -> Image.Image:
        payload = {
            **request.model_dump(exclude_none=True),
            **self.config.model_extra
        }
        response = self.session.post(
            self.config.api_base,
            json=payload,
            headers={"Authorization": f"Bearer {self.config.api_key}"},
            timeout=self.config.timeout
        )
        return self._process_response(response)

    def _process_response(self, response: requests.Response) -> Image.Image:
        response.raise_for_status()
        json_data = response.json()
        
        # Traverse JSON path
        image_url = json_data
        for key in self.config.response_json_path:
            if isinstance(image_url, list) and key.isdigit():
                image_url = image_url[int(key)]
            else:
                image_url = image_url.get(key)
        
        return self._download_image(image_url)

    def _download_image(self, url: str) -> Image.Image:
        response = self.session.get(url, timeout=self.config.timeout)
        response.raise_for_status()
        return Image.open(BytesIO(response.content))