# agents/image/agent.py
import hashlib
import logging
from cachetools import LRUCache
from PIL import Image
import orjson
from .manager import ModelManager
from schemas.image import GenerationRequest, BaseImageConfig

logger = logging.getLogger(__name__)

class ImageAgent:
    def __init__(
        self,
        main_config: BaseImageConfig,
        fallback_configs: list[BaseImageConfig] | None = None,
        cache_size: int = 100
    ):
        """
        Args:
            main_config: Основная конфигурация модели
            fallback_configs: Список конфигураций для фоллбэка
            cache_size: Размер кэша результатов
        """
        self.manager = ModelManager()
        self.main_model = self.manager.get_model(main_config)
        self.fallbacks = fallback_configs or []
        self.cache = LRUCache(maxsize=cache_size)

    def _generate_cache_key(self, request: GenerationRequest) -> str:
        """Генерирует ключ кэша на основе параметров запроса"""
        request_data = request.model_dump(
            exclude_none=True,
            exclude={"seed"}  # Исключаем seed для кэширования
        )
        return hashlib.sha256(
            orjson.dumps(request_data, option=orjson.OPT_SORT_KEYS)
        ).hexdigest()

    def generate(self, request: GenerationRequest) -> Image.Image:
        """Выполняет генерацию изображения с кэшированием и фоллбэком"""
        cache_key = self._generate_cache_key(request)
        
        if cache_key in self.cache:
            logger.debug("Returning cached result")
            return self.cache[cache_key]

        try:
            result = self.main_model.generate(request)
            self.cache[cache_key] = result
            return result
        except Exception as e:
            logger.error(f"Main model failed: {str(e)}")
            return self._try_fallbacks(request, cache_key)

    def _try_fallbacks(self, request: GenerationRequest, cache_key: str) -> Image.Image:
        """Пытается использовать фоллбэк-модели"""
        for fallback_config in self.fallbacks:
            try:
                model = self.manager.get_model(fallback_config)
                result = model.generate(request)
                self.cache[cache_key] = result
                logger.info(f"Fallback {type(fallback_config).__name__} succeeded")
                return result
            except Exception as e:
                logger.warning(f"Fallback failed: {str(e)}")
        
        raise RuntimeError("All generation attempts failed")

    def clear_cache(self):
        """Очищает кэш результатов"""
        self.cache.clear()