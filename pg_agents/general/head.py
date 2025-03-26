# agents/general/head.py
from agents import Agent

from . import prompts, tools


class HeadAgent(Agent):
    def __init__(self) -> None:
        super().__init__(
            name='HeadController',
            instructions=prompts.HEAD_PROMPT,
            tools=tools.HEAD_TOOLS,
            model='gpt-4-turbo',
        )
