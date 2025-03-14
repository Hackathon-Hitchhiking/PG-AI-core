from __future__ import annotations
from abc import ABC, abstractmethod
from typing import Any, Dict, Optional, Literal, List, Type, TypedDict
from pydantic import BaseModel, Field, field_validator
import logging
import torch
from langchain.chains.llm import LLMChain
from langchain_huggingface.llms import HuggingFacePipeline
from langchain_openai.llms import OpenAI
from llama_cpp import Llama
from llama_index.core import VectorStoreIndex
from llama_index.core import ServiceContext, PromptTemplate, Document
from transformers import AutoTokenizer, pipeline
from huggingface_hub import InferenceClient
import warnings
from vllm import SamplingParams, LLM

warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

TextBackend = Literal["hf", "api", "langchain", "llamacpp", "llamaindex", "vllm", "transformers"]

class TextModelConfig(BaseModel):
    backend: TextBackend = Field(..., description="Тип бэкенда для текстовой модели")
    model_name: Optional[str] = Field(None, min_length=1)
    api_base: Optional[str] = Field(None, min_length=3)
    api_key: Optional[str] = Field(None, min_length=1)
    model_path: Optional[str] = None
    device: str = Field(default="cuda" if torch.cuda.is_available() else "cpu")
    torch_dtype: Literal["auto", "float16", "float32"] = "auto"
    tokenizer_name: Optional[str] = None
    context_length: int = 4096
    temperature: float = 0.7
    max_new_tokens: int = 512
    top_p: float = 0.95
    langchain_template: Optional[str] = None
    llamacpp_params: Dict[str, Any] = Field(default_factory=dict)
    vector_store: Optional[str] = None
    quantized: bool = False
    use_safetensors: bool = True

    @field_validator("model_path")
    def validate_model_path(cls, v, values):
        backend = values.data.get("backend")
        if backend in ["llamacpp", "transformers"] and not v:
            raise ValueError("Model path is required for this backend")
        return v

class GenerationParams(TypedDict):
    temperature: float
    max_new_tokens: int
    top_p: float
    repetition_penalty: float
    stop_sequences: List[str]

class BaseTextModel(ABC):
    @abstractmethod
    def generate(self, prompt: str, params: GenerationParams) -> str:
        pass
    
    @classmethod
    @abstractmethod
    def from_config(cls, config: TextModelConfig) -> BaseTextModel:
        pass

class TextAgent:
    def __init__(
        self,
        config: TextModelConfig,
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
        self.config = config
        self.model = self._init_model()
        self.title_params = title_params
        self.content_params = content_params
        self.tokenizer = self._init_tokenizer()
        self.vector_index = self._init_vector_index()
        self.logger = logging.getLogger(self.__class__.__name__)

    def _init_model(self) -> BaseTextModel:
        backend_registry: Dict[TextBackend, Type[BaseTextModel]] = {
            "hf": HFTextModel,
            "api": APITextModel,
            "langchain": LangChainModel,
            "llamacpp": LlamaCppModel,
            "llamaindex": LlamaIndexModel,
            "vllm": vLLMModel,
            "transformers": transformersModel
        }
        return backend_registry[self.config.backend].from_config(self.config)

    def _init_tokenizer(self):
        try:
            return AutoTokenizer.from_pretrained(
                self.config.tokenizer_name or self.config.model_name,
                use_fast=True
            )
        except:
            return None

    def _init_vector_index(self):
        if self.config.vector_store and self.config.backend == "llamaindex":
            return VectorStoreIndex.from_vector_store(self.config.vector_store)
        return None

    def generate_title(self, context: str) -> str:
        augmented_prompt = f"Generate concise title for: {context}"
        return self._process_output(
            self.model.generate(augmented_prompt, self.title_params),
            self.title_params
        )

    def generate_content(self, context: str, format_hint: str = "paragraph") -> str:
        augmented_prompt = f"Generate detailed {format_hint} about: {context}"
        if self.vector_index:
            augmented_prompt += f"\nRelevant context: {self._get_related_context(context)}"
        return self._process_output(
            self.model.generate(augmented_prompt, self.content_params),
            self.content_params
        )

    def _get_related_context(self, query: str) -> str:
        return self.vector_index.query(query, similarity_top_k=3).response

    def _process_output(self, text: str, params: GenerationParams) -> str:
        if self.tokenizer:
            tokens = self.tokenizer.encode(text)
            text = self.tokenizer.decode(tokens[:params["max_new_tokens"]])
        return text.split(params["stop_sequences"][0])[0].strip()

class HFTextModel(BaseTextModel):
    def __init__(self, pipeline: Any):
        self.pipeline = pipeline

    @classmethod
    def from_config(cls, config: TextModelConfig) -> HFTextModel:
        pipe = pipeline(
            "text-generation",
            model=config.model_name,
            device=config.device,
            torch_dtype=config.torch_dtype,
            max_length=config.context_length
        )
        return cls(pipe)

    def generate(self, prompt: str, params: GenerationParams) -> str:
        output = self.pipeline(
            prompt,
            temperature=params["temperature"],
            max_new_tokens=params["max_new_tokens"],
            top_p=params["top_p"],
            repetition_penalty=params["repetition_penalty"],
            pad_token_id=self.pipeline.tokenizer.eos_token_id
        )
        return output[0]["generated_text"]

class APITextModel(BaseTextModel):
    def __init__(self, client: Any):
        self.client = client

    @classmethod
    def from_config(cls, config: TextModelConfig) -> APITextModel:
        if "openai" in config.api_base:
            return cls(OpenAI(api_key=config.api_key))
        return cls(InferenceClient(model=config.model_name, token=config.api_key))

    def generate(self, prompt: str, params: GenerationParams) -> str:
        if isinstance(self.client, OpenAI):
            return self.client.complete(prompt, **params).text
        return self.client.text_generation(
            prompt,
            temperature=params["temperature"],
            max_new_tokens=params["max_new_tokens"],
            top_p=params["top_p"],
            repetition_penalty=params["repetition_penalty"]
        )

class LangChainModel(BaseTextModel):
    def __init__(self, chain: LLMChain):
        self.chain = chain

    @classmethod
    def from_config(cls, config: TextModelConfig) -> LangChainModel:
        llm = HuggingFacePipeline.from_model_id(
            model_id=config.model_name,
            task="text-generation",
            device=config.device,
            pipeline_kwargs={
                "max_length": config.context_length,
                "temperature": config.temperature
            }
        )
        template = config.langchain_template or "{input}"
        return cls(LLMChain(llm=llm, prompt=PromptTemplate.from_template(template)))

    def generate(self, prompt: str, params: GenerationParams) -> str:
        return self.chain.run(
            input=prompt,
            temperature=params["temperature"],
            max_length=params["max_new_tokens"]
        )

class LlamaCppModel(BaseTextModel):
    def __init__(self, llm: Llama):
        self.llm = llm

    @classmethod
    def from_config(cls, config: TextModelConfig) -> LlamaCppModel:
        return cls(Llama(
            model_path=config.model_path,
            n_ctx=config.context_length,
            n_gpu_layers=-1 if torch.cuda.is_available() else 0,
            **config.llamacpp_params
        ))

    def generate(self, prompt: str, params: GenerationParams) -> str:
        output = self.llm(
            prompt,
            temperature=params["temperature"],
            max_tokens=params["max_new_tokens"],
            top_p=params["top_p"],
            repeat_penalty=params["repetition_penalty"],
            stop=params["stop_sequences"]
        )
        return output["choices"][0]["text"]

class LlamaIndexModel(BaseTextModel):
    def __init__(self, service_context: ServiceContext):
        self.service_context = service_context

    @classmethod
    def from_config(cls, config: TextModelConfig) -> LlamaIndexModel:
        llm = ServiceContext.from_defaults(llm=HuggingFacePipeline.from_model_id(
            model_id=config.model_name,
            task="text-generation",
            device=config.device
        ))
        return cls(ServiceContext.from_defaults(llm_predictor=llm))

    def generate(self, prompt: str, params: GenerationParams) -> str:
        index = VectorStoreIndex.from_documents([Document(prompt)])
        return index.query(
            prompt,
            service_context=self.service_context,
            similarity_top_k=3,
            response_mode="compact"
        ).response

class vLLMModel(BaseTextModel):
    def __init__(self, engine: Any):
        self.engine = engine

    @classmethod
    def from_config(cls, config: TextModelConfig) -> vLLMModel:
        return cls(LLM(
            model=config.model_path,
            tensor_parallel_size=1 if config.device == "cpu" else torch.cuda.device_count(),
            quantization="awq" if config.quantized else None
        ))

    def generate(self, prompt: str, params: GenerationParams) -> str:
        sampling_params = SamplingParams(
            temperature=params["temperature"],
            max_tokens=params["max_new_tokens"],
            top_p=params["top_p"],
            repetition_penalty=params["repetition_penalty"]
        )
        outputs = self.engine.generate([prompt], sampling_params)
        return outputs[0].outputs[0].text

class transformersModel(BaseTextModel):
    def __init__(self, model: Any):
        self.model = model

    @classmethod
    def from_config(cls, config: TextModelConfig) -> transformersModel:
        from transformers import AutoModelForCausalLM
        return cls(AutoModelForCausalLM.from_pretrained(
            config.model_path,
            model_type="llama",
            gpu_layers=50 if torch.cuda.is_available() else 0
        ))

    def generate(self, prompt: str, params: GenerationParams) -> str:
        return self.model(
            prompt,
            temperature=params["temperature"],
            max_new_tokens=params["max_new_tokens"],
            top_p=params["top_p"],
            repetition_penalty=params["repetition_penalty"]
        )