from agents import Agent

from . import prompts, tools


class ImageAgent(Agent):
    def __init__(self):
        super().__init__(
            name='ImageAgent',
            instructions=prompts.IMAGE_PROMPT,
            tools=tools.IMAGE_TOOLS,
        )
