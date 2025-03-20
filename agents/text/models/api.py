from typing import Any

import requests

from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import APIConfig, GenerationParams


@ModelRegistry.register(APIConfig)
class APIModel(BaseTextModel):
    def __init__(self, session: requests.Session, config: APIConfig) -> None:
        self.session = session
        self.config: APIConfig = config

    @classmethod
    def from_config(cls, config: APIConfig) -> 'APIModel':
        session = requests.Session()
        session.headers.update(config.headers)
        if config.api_key:
            session.headers['Authorization'] = f'Bearer {config.api_key}'
        return cls(session, config)

    def generate(self, prompt: str, params: GenerationParams) -> str:
        request_data = {'prompt': prompt, **params}

        response = self.session.post(self.config.api_base, json=request_data, timeout=self.config.timeout)
        response.raise_for_status()

        return self._process_response(response)

    def _process_response(self, response: requests.Response) -> str:
        if self.config.response_format == 'json':
            json_data = response.json()
            # Traverse JSON path to get the text response
            result: Any = json_data
            for key in self.config.response_path:
                result = result[key] if isinstance(result, list) and isinstance(key, int) else result.get(str(key))
            return str(result)
        return response.text
