from agents import Agent

from . import prompts, tools


class TextAgent(Agent):
    def __init__(self) -> None:
        super().__init__(name='TextAgent', instructions=prompts.TEXT_PROMPT, tools=tools.TEXT_TOOLS, model='gpt-4o')
