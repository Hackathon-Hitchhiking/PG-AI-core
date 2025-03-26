from agents import Agent

from . import prompts, tools


class SlideAgent(Agent):
    def __init__(self) -> None:
        super().__init__(name='SlideAgent', instructions=prompts.SLIDE_PROMPT, tools=tools.SLIDE_TOOLS, model='gpt-4o')
