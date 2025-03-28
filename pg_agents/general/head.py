from agents import Agent, ComputerTool, FileSearchTool, FunctionTool, WebSearchTool


class HeadAgent(Agent):
    def __init__(self, tools: list[FunctionTool | FileSearchTool | WebSearchTool | ComputerTool]) -> None:
        super().__init__(name='HeadController', instructions=tools, tools=tools, model='gpt-4o')
