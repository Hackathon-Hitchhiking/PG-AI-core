# agents/general/text.py
from agents import Agent

from . import prompts, tools


class TextAgent(Agent):
    def __init__(self):
        super().__init__(
            name='TextAgent',
            instructions=prompts.TEXT_PROMPT,
            tools=tools.TEXT_TOOLS,
            output_type=str,
        )
