from llama_cpp import Llama

from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import GenerationParams, LlamaCppConfig


@ModelRegistry.register(LlamaCppConfig)
class LlamaCppModel(BaseTextModel):
    def __init__(self, llm: Llama) -> None:
        self.llm = llm

    @classmethod
    def from_config(cls, config: LlamaCppConfig) -> 'LlamaCppModel':
        return cls(
            Llama(
                model_path=config.model_path,
                n_ctx=config.context_length,
                n_gpu_layers=config.n_gpu_layers,
                main_gpu=config.main_gpu,
                tensor_split=config.tensor_split,
                **config.llamacpp_params,
            )
        )

    def generate(self, prompt: str, params: GenerationParams) -> str:
        output = self.llm(
            prompt,
            temperature=params['temperature'],
            max_tokens=params['max_new_tokens'],
            top_p=params['top_p'],
            repeat_penalty=params['repetition_penalty'],
            stop=params['stop_sequences'],
        )
        return output['choices'][0]['text']
