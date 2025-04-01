from agents import Agent

from . import prompts


class ChartAgent(Agent):
    def __init__(self, tools: list) -> None:
        super().__init__(name='ChartAgent', instructions=prompts.CHART_PROMPT, tools=tools, model='gpt-4o')
