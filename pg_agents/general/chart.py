# agents/general/chart.py
from agents import Agent

from . import prompts, tools


class ChartAgent(Agent):
    def __init__(self):
        super().__init__(
            name='ChartAgent',
            instructions=prompts.CHART_PROMPT,
            tools=tools.CHART_TOOLS,
            # output_type=dict
        )
