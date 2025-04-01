from agents import Agent, ComputerTool, FileSearchTool, FunctionTool, WebSearchTool

from pg_agents.general import prompts


class HeadAgent(Agent):
    def __init__(self, tools: list[FunctionTool | FileSearchTool | WebSearchTool | ComputerTool]) -> None:
        super().__init__(name='HeadController', instructions=prompts.HEAD_PROMPT, tools=tools, model='gpt-4o')
