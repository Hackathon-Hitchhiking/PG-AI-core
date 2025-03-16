from typing import Dict
from agents.text.schemas import BaseTextConfig
from agents.text.models.base import BaseTextModel, ModelRegistry

class ModelManager:
    def __init__(self):
        self._instances: Dict[str, BaseTextModel] = {}

    def get_model(self, config: BaseTextConfig) -> BaseTextModel:
        """Get or create model instance for given configuration"""
        config_key = str(config.model_dump())
        
        if config_key not in self._instances:
            model_class = ModelRegistry.get_model_class(config)
            self._instances[config_key] = model_class.from_config(config)
            
        return self._instances[config_key]

    def clear(self):
        """Clear all model instances"""
        self._instances.clear()
