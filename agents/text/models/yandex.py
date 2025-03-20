import logging

from yandex_cloud_ml_sdk import YCloudML
from yandex_cloud_ml_sdk._models.completions.model import GPTModel

from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import GenerationParams, YandexConfig


logger = logging.getLogger(__name__)


@ModelRegistry.register(YandexConfig)
class YandexModel(BaseTextModel):
    def __init__(self, model: GPTModel, config: YandexConfig) -> None:
        self.model = model
        self.config = config

    @classmethod
    def from_config(cls, config: YandexConfig) -> 'YandexModel':
        sdk = YCloudML(folder_id=config.folder_id, auth=config.api_key)
        model = sdk.models.completions('yandexgpt')
        return cls(model, config)

    def generate(self, prompt: str, params: GenerationParams) -> str:
        try:
            messages = [{'role': 'user', 'text': prompt}]

            config_params = {}
            if 'max_new_tokens' in params:
                config_params['max_tokens'] = params['max_new_tokens']
            if 'temperature' in params:
                config_params['temperature'] = params['temperature']

            configured_model = self.model.configure(**config_params)
            result = configured_model.run(messages=messages)

            # Get text directly from the alternative
            return result.alternatives[0].text if result and result.alternatives else ''

        except Exception as e:
            logger.exception(f'Yandex generation failed: {str(e)}')
            msg = 'Text generation failed'
            raise RuntimeError(msg) from e
