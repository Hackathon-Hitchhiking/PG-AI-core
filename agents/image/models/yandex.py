from yandex_cloud_ml_sdk import YCloudML
from PIL import Image
from io import BytesIO
import logging
from schemas.image import YandexConfig, GenerationRequest
from .base import BaseImageModel, ModelRegistry

from yandex_cloud_ml_sdk._models.image_generation.model import ImageGenerationModel

logger = logging.getLogger(__name__)

@ModelRegistry.register(YandexConfig)
class YandexModel(BaseImageModel):
    BASE_SIZE = 512

    def __init__(self, model: ImageGenerationModel, config: YandexConfig):
        self.model: ImageGenerationModel = model
        self.config: YandexConfig = config

    @classmethod
    def from_config(cls, config: YandexConfig) -> "YandexModel":
        sdk = YCloudML(
            folder_id=config.folder_id,
            auth=config.iam_token.get_secret_value()
        )
        return cls(sdk.models.image_generation("yandex-art"), config)

    def _calculate_ratios(self, width: int, height: int) -> tuple[int, int]:
        """Рассчитывает соотношения размеров с округлением до ближайшего допустимого значения."""
        width_ratio = max(1, round(width / self.BASE_SIZE))
        height_ratio = max(1, round(height / self.BASE_SIZE))
        
        if width_ratio * self.BASE_SIZE != width or height_ratio * self.BASE_SIZE != height:
            logger.warning(
                f"Adjusting size from {width}x{height} to "
                f"{width_ratio*self.BASE_SIZE}x{height_ratio*self.BASE_SIZE}"
            )
        return width_ratio, height_ratio

    def generate(self, request: GenerationRequest) -> Image.Image:
        width_ratio, height_ratio = self._calculate_ratios(request.width, request.height)

        configured_model: ImageGenerationModel = self.model.configure(
            # negative_prompt=request.negative_prompt or "", хз куда его передавать
            width_ratio=width_ratio,
            height_ratio=height_ratio,
            seed=request.seed or 42,
            # num_inference_steps=request.num_inference_steps,
            # guidance_scale=request.guidance_scale
        )

        try:
            operation = configured_model.run_deferred(request.prompt)
            result = operation.wait()
            
            if not result.image_bytes:
                raise ValueError("Empty response from Yandex API")
                
            return Image.open(BytesIO(result.image_bytes)).convert("RGB")
            
        except Exception as e:
            logger.error(f"Yandex generation failed: {str(e)}")
            raise RuntimeError("Image generation failed") from e