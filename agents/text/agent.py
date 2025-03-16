from __future__ import annotations
import hashlib
import logging
import warnings
from cachetools import LRUCache
import orjson
from agents.text.manager import ModelManager
from agents.text.schemas import GenerationParams, BaseTextConfig

warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

class TextAgent:
    def __init__(
        self,
        main_config: BaseTextConfig,
        fallback_configs: list[BaseTextConfig] | None = None,
        cache_size: int = 100,
        title_params: GenerationParams = {
            "temperature": 0.3,
            "max_new_tokens": 50,
            "top_p": 0.9,
            "repetition_penalty": 1.2,
            "stop_sequences": ["\n"]
        },
        content_params: GenerationParams = {
            "temperature": 0.7,
            "max_new_tokens": 500,
            "top_p": 0.95,
            "repetition_penalty": 1.1,
            "stop_sequences": ["\n\n"]
        }
    ):
        self.manager = ModelManager()
        self.main_model = self.manager.get_model(main_config)
        self.fallbacks = fallback_configs or []
        self.cache = LRUCache(maxsize=cache_size)
        self.title_params = title_params
        self.content_params = content_params

    def _generate_cache_key(self, prompt: str, params: GenerationParams) -> str:
        request_data = {
            "prompt": prompt,
            **{k: v for k, v in params.items() if k != "seed"}
        }
        return hashlib.sha256(
            orjson.dumps(request_data, option=orjson.OPT_SORT_KEYS)
        ).hexdigest()

    def generate(self, prompt: str, params: GenerationParams) -> str:
        """Генерирует текст с использованием основной или запасной модели"""
        cache_key = self._generate_cache_key(prompt, params)
        
        if cache_key in self.cache:
            logger.debug("Returning cached result")
            return self.cache[cache_key]

        try:
            result = self.main_model.generate(prompt, params)
            self.cache[cache_key] = result
            return result
        except Exception as e:
            logger.error(f"Main model failed: {str(e)}")
            return self._try_fallbacks(prompt, params, cache_key)

    def _try_fallbacks(self, prompt: str, params: GenerationParams, cache_key: str) -> str:
        """Пытается использовать запасные модели при сбое основной"""
        for fallback_config in self.fallbacks:
            try:
                model = self.manager.get_model(fallback_config)
                result = model.generate(prompt, params)
                self.cache[cache_key] = result
                logger.info(f"Fallback {type(fallback_config).__name__} succeeded")
                return result
            except Exception as e:
                logger.warning(f"Fallback failed: {str(e)}")
        
        raise RuntimeError("All generation attempts failed")

    def generate_title(self, context: str) -> str:
        """Генерирует заголовок для заданного контекста"""
        augmented_prompt = f"Generate concise title for: {context}"
        return self.generate(augmented_prompt, self.title_params)

    def generate_content(self, context: str, format_hint: str = "paragraph") -> str:
        """Генерирует контент с указанным форматом"""
        augmented_prompt = f"Generate detailed {format_hint} about: {context}"
        return self.generate(augmented_prompt, self.content_params)

    def clear_cache(self):
        """Очищает кэш результатов"""
        self.cache.clear()