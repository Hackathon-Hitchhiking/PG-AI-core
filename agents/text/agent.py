from __future__ import annotations

import hashlib
import logging
import warnings

import orjson

from cachetools import LRUCache

from agents.text.manager import ModelManager
from agents.text.schemas import BaseTextConfig, ContentStyle, ContentTone, GenerationParams
from utils.prompt_library import PromptLibrary


warnings.filterwarnings('ignore')
logger = logging.getLogger(__name__)


class TextAgent:
    def __init__(
        self,
        main_config: BaseTextConfig,
        fallback_configs: list[BaseTextConfig] | None = None,
        cache_size: int = 100,
    ) -> None:
        self.manager = ModelManager()
        self.main_model = self.manager.get_model(main_config)
        self.fallbacks = fallback_configs or []
        self.cache = LRUCache(maxsize=cache_size)
        self.prompt_library = PromptLibrary()

    def _generate_cache_key(self, prompt: str, params: GenerationParams) -> str:
        request_data = {'prompt': prompt, **{k: v for k, v in params.items() if k != 'seed'}}
        return hashlib.sha256(orjson.dumps(request_data, option=orjson.OPT_SORT_KEYS)).hexdigest()

    def generate(self, prompt: str, params: GenerationParams) -> str:
        """Генерирует текст с использованием основной или запасной модели"""
        cache_key = self._generate_cache_key(prompt, params)

        if cache_key in self.cache:
            logger.debug('Returning cached result')
            return self.cache[cache_key]

        try:
            result = self.main_model.generate(prompt, params)
            self.cache[cache_key] = result
            return result
        except Exception as e:
            logger.exception(f'Main model failed: {str(e)}')
            return self._try_fallbacks(prompt, params, cache_key)

    def _try_fallbacks(self, prompt: str, params: GenerationParams, cache_key: str) -> str:
        """Пытается использовать запасные модели при сбое основной"""
        for fallback_config in self.fallbacks:
            try:
                model = self.manager.get_model(fallback_config)
                result = model.generate(prompt, params)
                self.cache[cache_key] = result
                logger.info(f'Fallback {type(fallback_config).__name__} succeeded')
                return result
            except Exception as e:
                logger.warning(f'Fallback failed: {str(e)}')

        msg = 'All generation attempts failed'
        raise RuntimeError(msg)

    def generate_title(self, context: str, params: GenerationParams | None = None) -> str:
        """Генерирует заголовок с учетом стиля и тона"""
        base_params: GenerationParams = {
            'temperature': 0.3,
            'max_new_tokens': 50,
            'top_p': 0.9,
            'repetition_penalty': 1.2,
            'stop_sequences': ['\n'],
            'style': ContentStyle.ARTICLE,
            'tone': ContentTone.PROFESSIONAL,
        }
        if params:
            base_params.update(params)

        prompt = self.prompt_library.get_prompt(
            'title', context=context, style=base_params['style'], tone=base_params['tone']
        )

        return self.generate(prompt, base_params)

    def generate_content(self, context: str, params: GenerationParams | None = None, format_hint: str = 'текст') -> str:
        """Генерирует контент с учетом стиля, тона и формата"""
        base_params: GenerationParams = {
            'temperature': 0.7,
            'max_new_tokens': 500,
            'top_p': 0.95,
            'repetition_penalty': 1.1,
            'stop_sequences': ['\n\n'],
            'style': ContentStyle.ARTICLE,
            'tone': ContentTone.PROFESSIONAL,
            'length': 'medium',
        }
        if params:
            base_params.update(params)

        if not params or 'max_new_tokens' not in params:
            length = base_params.get('length', 'medium')
            length_tokens = {'short': 150, 'medium': 500, 'long': 1000}
            base_params['max_new_tokens'] = length_tokens[length]

        prompt = self.prompt_library.get_prompt(
            'content', context=context, style=base_params['style'], tone=base_params['tone'], format_hint=format_hint
        )

        return self.generate(prompt, base_params)

    def clear_cache(self) -> None:
        """Очищает кэш результатов"""
        self.cache.clear()
