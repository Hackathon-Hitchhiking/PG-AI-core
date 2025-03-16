# agents/image/models/api.py
import requests
from PIL import Image
from io import BytesIO
from agents.image.schemas import APIConfig, GenerationRequest
from agents.image.models.base import BaseImageModel, ModelRegistry

@ModelRegistry.register(APIConfig)
class APIModel(BaseImageModel):
    def __init__(self, session: requests.Session, config: APIConfig):
        self.session = session
        self.config: APIConfig = config

    @classmethod
    def from_config(cls, config: APIConfig) -> "APIModel":
        session = requests.Session()
        session.headers.update(config.headers)
        if config.api_key:
            session.headers["Authorization"] = f"Bearer {config.api_key.get_secret_value()}"
        return cls(session, config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        if self.config.method == "GET":
            return self._handle_get(request)
        return self._handle_post(request)

    def _handle_get(self, request: GenerationRequest) -> Image.Image:
        response = self.session.get(
            self.config.api_base,
            params={"prompt": request.prompt},
            timeout=self.config.timeout
        )
        return self._process_response(response)

    def _handle_post(self, request: GenerationRequest) -> Image.Image:
        response = self.session.post(
            self.config.api_base,
            json=request.model_dump(),
            timeout=self.config.timeout
        )
        return self._process_response(response)

    def _process_response(self, response: requests.Response) -> Image.Image:
        response.raise_for_status()
        json_data = response.json()
        
        # Traverse JSON path
        result = json_data
        for key in self.config.response_json_path:
            if isinstance(result, list) and isinstance(key, int):
                result = result[key]
            else:
                result = result.get(str(key))
        
        return self._download_image(result)

    def _download_image(self, url: str) -> Image.Image:
        response = self.session.get(url, timeout=self.config.timeout)
        response.raise_for_status()
        return Image.open(BytesIO(response.content))