import getpass
import os

import httpx

from agno.agent import Agent
from agno.models.openai import OpenAIChat
from agno.team import Team
from dotenv import load_dotenv
from openai import AsyncOpenAI, OpenAI

from pptx_manager.main import PPTXManager


load_dotenv()

if not os.environ.get('OPENAI_API_KEY'):
    os.environ['OPENAI_API_KEY'] = getpass.getpass('Enter API key for OpenAI: ')


http_async_client = httpx.AsyncClient(proxy='http://127.0.0.1:1080')
http_sync_client = httpx.Client(proxy='http://127.0.0.1:1080')

open_async_client = AsyncOpenAI(
    http_client=http_async_client,
)

open_sync_client = OpenAI(
    http_client=http_sync_client,
)

pr = PPTXManager('../test_data/test_dit.pptx')

image_json = pr.get_all_text_frame_json()

pr_json = pr.get_all_text_frame_json()

model = OpenAIChat(id='gpt-4o-mini', client=open_sync_client, async_client=open_async_client)

text_agent = Agent(
    name='Text Agent',
    role='Changing the texts on the presentation',
    instructions=[
        'You are the text agent that should work with the text on the presentation',
        'There is the two main entities that you should know slide_id and the shape_id that means the number of the slide and the number of the shape.',
    ],
    tools=[pr.update_text_frame_shape],
    model=model,
    show_tool_calls=True,
    markdown=True,
)

image_agent = Agent(
    name='Image Agent',
    role='Changing the images on the presentation',
    instructions='',
    tools=[],
    model=model,
    show_tool_calls=True,
    markdown=True,
)

head_agent = Team(
    mode='route',
    members=[text_agent],
    model=model,
    success_criteria='you solve all tasks that giving you user',
    instructions=[
        'You are the manager of the sub agents that can change the presentation',
        ' There is the two main entities that you should know slide_id and the shape_id that means the number of the slide and the number of the shape.',
    ],
    show_tool_calls=True,
    markdown=True,
)

head_agent.run("change the size to the 40 on the first slide to the 'TEST'")

pr.save('test.pptx')
