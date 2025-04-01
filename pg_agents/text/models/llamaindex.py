from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import GenerationParams, LlamaIndexConfig
from langchain_huggingface import HuggingFacePipeline
from llama_index.core import Document, ServiceContext, VectorStoreIndex


@ModelRegistry.register(LlamaIndexConfig)
class LlamaIndexModel(BaseTextModel):
    def __init__(self, service_context: ServiceContext, config: LlamaIndexConfig) -> None:
        self.service_context = service_context
        self.config = config
        self.index = None

    @classmethod
    def from_config(cls, config: LlamaIndexConfig) -> 'LlamaIndexModel':
        llm = HuggingFacePipeline.from_model_id(
            model_id=config.model_name, task='text-generation', device=config.device
        )
        service_context = ServiceContext.from_defaults(
            llm=llm, chunk_size=config.chunk_size, chunk_overlap=config.chunk_overlap
        )
        return cls(service_context, config)

    def generate(self, prompt: str, params: GenerationParams) -> str:
        if not self.index:
            self.index = VectorStoreIndex.from_documents([Document(text=prompt)], service_context=self.service_context)

        response = self.index.query(
            prompt, similarity_top_k=self.config.similarity_top_k, response_mode=self.config.response_mode, **params
        )
        return response.response
