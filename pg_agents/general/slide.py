from agents import Agent

from . import prompts


class SlideAgent(Agent):
    def __init__(self, tools: list) -> None:
        super().__init__(name='SlideAgent', instructions=prompts.SLIDE_PROMPT, tools=tools, model='gpt-4o')
