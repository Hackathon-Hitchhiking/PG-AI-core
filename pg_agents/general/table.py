from agents import Agent

from . import prompts, tools


class TableAgent(Agent):
    def __init__(self) -> None:
        super().__init__(name='TableAgent', instructions=prompts.TABLE_PROMPT, tools=tools.TABLE_TOOLS, model='gpt-4o')
