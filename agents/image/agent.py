import hashlib
import logging

from pathlib import Path

import orjson

from cachetools import LRUCache
from PIL import Image
from pydantic import ValidationError

from agents.image.exceptions import ImageGenerationError
from agents.image.manager import ModelManager
from agents.image.schemas import BaseImageConfig, GenerationRequest


logger = logging.getLogger(__name__)

class ImageAgent:
    def __init__(
        self,
        main_config: BaseImageConfig,
        fallback_configs: list[BaseImageConfig] | None = None,
        cache_size: int = 100,
        cache_dir: Path | None = None
    ) -> None:
        """Initialize ImageAgent with configs and caching.

        Args:
            main_config: Primary model configuration
            fallback_configs: List of fallback model configurations
            cache_size: Max number of results to cache in memory
            cache_dir: Optional directory for disk caching
        """
        self.manager = ModelManager()
        try:
            self.main_model = self.manager.get_model(main_config)
        except Exception as e:
            logger.exception("Failed to initialize main model")
            msg = "Main model initialization failed"
            raise ImageGenerationError(msg) from e

        self.fallbacks = []
        if fallback_configs:
            try:
                for config in fallback_configs:
                    self.fallbacks.append(self.manager.get_model(config))
            except Exception as e:
                logger.warning(f"Failed to initialize fallback model: {e}")

        self.cache = LRUCache(maxsize=cache_size)
        self.cache_dir = cache_dir

    def _generate_cache_key(self, request: GenerationRequest) -> str:
        """Генерирует ключ кэша на основе параметров запроса"""
        request_data = request.model_dump(
            exclude_none=True,
            exclude={"seed"}  # Исключаем seed для кэширования
        )
        return hashlib.sha256(
            orjson.dumps(request_data, option=orjson.OPT_SORT_KEYS)
        ).hexdigest()

    def generate(self, request: GenerationRequest, retry_count: int = 2) -> Image.Image:
        """Generate image from request with retries and fallbacks.

        Args:
            request: Image generation parameters
            retry_count: Number of retries for transient failures

        Returns:
            Generated PIL Image

        Raises:
            ImageGenerationError: If generation fails after retries and fallbacks
        """
        try:
            request.validate()
        except ValidationError as e:
            logger.exception("Invalid generation request")
            msg = "Invalid request parameters"
            raise ImageGenerationError(msg) from e

        cache_key = self._generate_cache_key(request)
        if cache_key in self.cache:
            logger.debug("Returning cached result")
            return self.cache[cache_key]

        attempt = 0
        while attempt < retry_count:
            image = self.main_model.generate(request)
            self.cache[cache_key] = image
            return image
        attempt += 1

        # Try fallbacks
        try:
            for fallback in self.fallbacks:
                logger.info(f"Trying fallback model: {fallback.__class__.__name__}")
                image = fallback.generate(request)
                self.cache[cache_key] = image
                return image
        except Exception as e:
            logger.warning(f"All fallback models failed: {e}")

        msg = "All generation attempts failed"
        raise ImageGenerationError(msg)

    def clear_cache(self) -> None:
        """Очищает кэш результатов"""
        self.cache.clear()
