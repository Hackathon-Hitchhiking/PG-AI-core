import hashlib

import orjson

from agents.image.models import BaseImageModel, ModelRegistry
from agents.image.schemas import BaseImageConfig
from cachetools import LRUCache


class ModelManager:
    def __init__(self, max_size: int = 10) -> None:
        self.cache = LRUCache(maxsize=max_size)

    def _get_config_hash(self, config: BaseImageConfig) -> str:
        """Генерирует уникальный хэш для конфигурации"""
        config_dict = config.model_dump(mode='json')
        return hashlib.sha256(orjson.dumps(config_dict, option=orjson.OPT_SORT_KEYS)).hexdigest()

    def get_model(self, config: BaseImageConfig) -> BaseImageModel:
        """Возвращает модель по конфигурации, используя кэш"""
        config_hash = self._get_config_hash(config)

        if config_hash not in self.cache:
            model_class = ModelRegistry.get_model_class(config)
            self.cache[config_hash] = model_class.from_config(config)

        return self.cache[config_hash]
