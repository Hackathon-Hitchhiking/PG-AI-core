from yandex_cloud_ml_sdk import YCloudML
from PIL import Image
from schemas.image import YandexConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry, ModelType

@ModelRegistry.register(ModelType.YANDEX)
class YandexModel(BaseImageModel):
    def __init__(self, config: YandexConfig):
        self.config = config
        self.client = YCloudML(
            folder_id=config.folder_id,
            iam_token=config.iam_token.get_secret_value()
        ).models.yandex_art(preview=config.use_preview)

    @classmethod
    def from_config(cls, config: YandexConfig) -> "YandexModel":
        return cls(config)

    def generate(self, request: GenerationRequest) -> Image.Image:
        result = self.client.run(
            request.prompt,
            negative_prompt=request.negative_prompt,
            width=request.width,
            height=request.height,
            num_steps=request.num_inference_steps,
            guidance_scale=request.guidance_scale,
            seed=request.seed or 0
        )
        return result[0].to_pil_image()