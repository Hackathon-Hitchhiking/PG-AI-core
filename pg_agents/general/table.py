from agents import Agent

from . import prompts


class TableAgent(Agent):
    def __init__(self, tools: list) -> None:
        super().__init__(name='TableAgent', instructions=prompts.TABLE_PROMPT, tools=tools, model='gpt-4o')
