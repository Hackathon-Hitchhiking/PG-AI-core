from agents import Agent

from . import prompts


class ImageAgent(Agent):
    def __init__(self, tools: list) -> None:
        super().__init__(name='ImageAgent', instructions=prompts.IMAGE_PROMPT, tools=tools, model='gpt-4o')
