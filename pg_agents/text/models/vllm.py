from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import GenerationParams, vLLMConfig
from vllm import LLM, SamplingParams


@ModelRegistry.register(vLLMConfig)
class VLLMModel(BaseTextModel):
    def __init__(self, engine: LLM) -> None:
        self.engine = engine

    @classmethod
    def from_config(cls, config: vLLMConfig) -> 'VLLMModel':
        tensor_parallel = 1 if config.device == 'mps' else config.tensor_parallel_size

        return cls(
            LLM(
                model=config.model_path,
                tensor_parallel_size=tensor_parallel,
                gpu_memory_utilization=config.gpu_memory_utilization,
                max_model_len=config.max_model_len or config.context_length,
                enforce_eager=config.enforce_eager,
                trust_remote_code=True,
                quantization='awq' if config.quantized else None,
            )
        )

    def generate(self, prompt: str, params: GenerationParams) -> str:
        sampling_params = SamplingParams(
            temperature=params['temperature'],
            max_tokens=params['max_new_tokens'],
            top_p=params['top_p'],
            top_k=params['top_k'],
            presence_penalty=params['presence_penalty'],
            frequency_penalty=params['frequency_penalty'],
            repetition_penalty=params['repetition_penalty'],
            stop=params['stop_sequences'],
        )

        outputs = self.engine.generate([prompt], sampling_params)
        return outputs[0].outputs[0].text
