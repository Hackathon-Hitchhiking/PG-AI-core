# agents/image/agent.py
import hashlib
import logging
from PIL import Image
from cachetools import LRUCache
import orjson
from pydantic import BaseModel
from schemas.image import GenerationRequest, ImageModelConfig
from agents.image.manager import ModelManager

logger = logging.getLogger(__name__)

class FallbackStrategy(BaseModel):
    config: ImageModelConfig
    priority: int = 0

class ImageAgent:
    def __init__(
        self,
        main_config: ImageModelConfig,
        fallbacks: list[FallbackStrategy] | None = None,
        max_cache_size: int = 100
    ):
        self.main_config = main_config
        self.fallbacks = sorted(
            fallbacks or [], 
            key=lambda x: x.priority, 
            reverse=True
        )
        self.manager = ModelManager(max_cache_size=max_cache_size)
        self.request_cache = LRUCache(maxsize=max_cache_size)

    def _request_key(self, request: GenerationRequest) -> str:
        return hashlib.sha256(
            orjson.dumps(request.model_dump(exclude={"seed"}))
        ).hexdigest()

    def generate(self, request: GenerationRequest) -> Image.Image:
        cache_key = self._request_key(request)
        
        if cache_key in self.request_cache:
            return self.request_cache[cache_key]

        try:
            model = self.manager.get_model(self.main_config)
            image = model.generate(request)
            self.request_cache[cache_key] = image
            return image
        except Exception as e:
            logger.error(f"Main model failed: {e}")
            return self._handle_fallback(request, cache_key)

    def _handle_fallback(self, request: GenerationRequest, cache_key: str) -> Image.Image:
        for strategy in self.fallbacks:
            try:
                model = self.manager.get_model(strategy.config)
                image = model.generate(request)
                self.request_cache[cache_key] = image
                return image
            except Exception as e:
                logger.warning(f"Fallback {strategy.config.model_type} failed: {e}")
        
        raise RuntimeError("All image generation attempts failed")