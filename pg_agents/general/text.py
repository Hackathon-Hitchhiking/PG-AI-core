from agents import Agent

from . import prompts


class TextAgent(Agent):
    def __init__(self, tools: list) -> None:
        super().__init__(name='TextAgent', instructions=prompts.TEXT_PROMPT, tools=tools, model='gpt-4o')
