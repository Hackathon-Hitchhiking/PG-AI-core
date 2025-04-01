from typing import Any

from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import GenerationParams, TransformersConfig
from transformers import AutoModelForCausalLM, AutoTokenizer


@ModelRegistry.register(TransformersConfig)
class TransformersModel(BaseTextModel):
    def __init__(self, model: Any, tokenizer: AutoTokenizer) -> None:
        self.model: AutoModelForCausalLM = model
        self.tokenizer: AutoTokenizer = tokenizer

    @classmethod
    def from_config(cls, config: TransformersConfig) -> 'TransformersModel':
        model_kwargs = {
            'device_map': config.device,
            'trust_remote_code': config.trust_remote_code,
            'use_auth_token': config.use_auth_token,
            'load_in_8bit': config.load_in_8bit,
        }

        if config.use_safetensors:
            model_kwargs['use_safetensors'] = True

        model = AutoModelForCausalLM.from_pretrained(config.model_path, **model_kwargs)

        tokenizer = AutoTokenizer.from_pretrained(
            config.tokenizer_name or config.model_path,
            trust_remote_code=config.trust_remote_code,
            use_auth_token=config.use_auth_token,
        )

        return cls(model, tokenizer)

    def generate(self, prompt: str, params: GenerationParams) -> str:
        inputs = self.tokenizer(prompt, return_tensors='pt')
        inputs = inputs.to(self.model.device)

        generation_kwargs = {
            'max_new_tokens': params['max_new_tokens'],
            'temperature': params['temperature'],
            'top_p': params['top_p'],
            'top_k': params['top_k'],
            'repetition_penalty': params['repetition_penalty'],
            'do_sample': params['do_sample'],
            'num_beams': params['num_beams'],
            'pad_token_id': self.tokenizer.pad_token_id,
        }

        outputs = self.model.generate(**inputs, **generation_kwargs)
        return self.tokenizer.decode(outputs[0], skip_special_tokens=params['skip_special_tokens'])
