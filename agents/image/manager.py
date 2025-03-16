# agents/image/manager.py
import hashlib
import orjson
from cachetools import LRUCache
from agents.image.base import BaseImageModel, ModelRegistry
from schemas.image import ImageModelConfig

class ModelManager:
    def __init__(self, max_cache_size: int = 10):
        self.cache = LRUCache(maxsize=max_cache_size)

    def _config_key(self, config: ImageModelConfig) -> str:
        config_data = config.model_dump(mode="json")
        return hashlib.sha256(
            orjson.dumps(config_data, option=orjson.OPT_SORT_KEYS)
        ).hexdigest()

    def get_model(self, config: ImageModelConfig) -> BaseImageModel:
        key = self._config_key(config)
        if key not in self.cache:
            model_cls = ModelRegistry.get_model_class(config.model_type)
            self.cache[key] = model_cls.from_config(config)
        return self.cache[key]