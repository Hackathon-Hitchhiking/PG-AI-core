from langchain.chains.llm import LLMChain
from langchain.prompts import PromptTemplate
from langchain_huggingface.llms import HuggingFacePipeline

from agents.text.models.base import BaseTextModel, ModelRegistry
from agents.text.schemas import GenerationParams, LangChainConfig


@ModelRegistry.register(LangChainConfig)
class LangChainModel(BaseTextModel):
    def __init__(self, chain: LLMChain) -> None:
        self.chain = chain

    @classmethod
    def from_config(cls, config: LangChainConfig) -> "LangChainModel":
        llm = HuggingFacePipeline.from_model_id(
            model_id=config.model_name,
            task="text-generation",
            device=config.device,
            pipeline_kwargs={
                "max_length": config.context_length,
                "temperature": 0.7
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
