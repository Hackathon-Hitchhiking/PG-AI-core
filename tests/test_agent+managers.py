import asyncio
import os

import httpx

from agents import (
    Agent,
    HandoffOutputItem,
    ItemHelpers,
    MessageOutputItem,
    Runner,
    ToolCallItem,
    ToolCallOutputItem,
    TResponseInputItem,
    set_default_openai_client,
)
from dotenv import load_dotenv
from openai import AsyncOpenAI

from pg_agents.general import prompts
from pptx_manager.main import PPTXManager


load_dotenv()

if not os.environ.get('OPENAI_API_KEY'):
    msg = 'добавь ключ в .env'
    raise RuntimeError(msg)

pr = PPTXManager('../test_data/test_dit.pptx')

image_json = pr.get_all_text_frame_json()

pr_json = pr.get_all_text_frame_json()

http_async_client = httpx.AsyncClient(proxy='http://127.0.0.1:1080')

custom_client = AsyncOpenAI(
    http_client=http_async_client,
)

set_default_openai_client(custom_client, True)

text_agent = Agent(
    name='TextAgent', instructions=prompts.TEXT_PROMPT.format(image_json), tools=[pr.update_text_frame_shape]
)

tools = [
    text_agent.as_tool(
        tool_name='text_manager',
        tool_description='Управляет текстовыми блоками: создание, редактирование, стилизация текста на слайдах.',
    ),
]

head_agent = Agent(name='HeadController', instructions=prompts.HEAD_PROMPT.format(pr_json), tools=tools, model='gpt-4o')


async def main():
    input_items: list[TResponseInputItem] = []

    while True:
        user_input = input('Enter your message: ')
        input_items.append({'content': user_input, 'role': 'user'})
        result = await Runner.run(head_agent, input_items)

        for new_item in result.new_items:
            agent_name = new_item.agent.name
            if isinstance(new_item, MessageOutputItem):
                print(f'{agent_name}: {ItemHelpers.text_message_output(new_item)}')
            elif isinstance(new_item, HandoffOutputItem):
                print(f'Handed off from {new_item.source_agent.name} to {new_item.target_agent.name}')
            elif isinstance(new_item, ToolCallItem):
                print(f'{agent_name}: Calling a tool {new_item.raw_item.name}')
            elif isinstance(new_item, ToolCallOutputItem):
                print(f'{agent_name}: Tool call output: {new_item.output}')
            else:
                print(f'{agent_name}: Skipping item: {new_item.__class__.__name__}')
        input_items = result.to_input_list()


if __name__ == '__main__':
    asyncio.run(main())
