from yandex_cloud_ml_sdk import YCloudML
from yandex_cloud_ml_sdk._models.completions.model import GPTModel
import logging
from agents.text.schemas import YandexConfig, GenerationParams
from agents.text.models.base import BaseTextModel, ModelRegistry

logger = logging.getLogger(__name__)

@ModelRegistry.register(YandexConfig)
class YandexModel(BaseTextModel):
    def __init__(self, model: GPTModel, config: YandexConfig):
        self.model = model
        self.config = config

    @classmethod
    def from_config(cls, config: YandexConfig) -> "YandexModel":
        sdk = YCloudML(folder_id=config.folder_id, auth=config.api_key)
        model = sdk.models.completions('yandexgpt')
        return cls(model, config)

    def generate(self, prompt: str, params: GenerationParams) -> str:
        try:
            generation_params = {
                "temperature": params["temperature"],
                "max_tokens": params["max_new_tokens"],
                "stream": False
            }
            
            result = self.model.run(prompt, **generation_params)
            return str(result[0]) if result else ""
            
        except Exception as e:
            logger.error(f"Yandex generation failed: {str(e)}")
            raise RuntimeError("Text generation failed") from e
