import getpass
import json
import os
import signal

from pathlib import Path

import httpx

from dotenv import load_dotenv
from huggingface_hub.inference._generated.types.chat_completion import (
    ChatCompletionOutputFunctionDefinition,
    ChatCompletionOutputToolCall,
)
from langchain.chat_models import init_chat_model
from langchain_community.chat_message_histories import ChatMessageHistory
from langchain_core.tools import tool

from tools.pptx_tools import PPTXManager


load_dotenv()

if not os.environ.get('OPENAI_API_KEY'):
    os.environ['OPENAI_API_KEY'] = getpass.getpass('Enter API key for OpenAI: ')

manager = PPTXManager('test_sources/test_dit.pptx')
parse_docstring = True
tools = [
    tool(manager.get_presentation_info, parse_docstring=parse_docstring),
    tool(manager.get_slide_details, parse_docstring=parse_docstring),
    tool(manager.analyze_slide_design, parse_docstring=parse_docstring),
    tool(manager.create_new_slide, parse_docstring=parse_docstring),
    tool(manager.duplicate_slide, parse_docstring=parse_docstring),
    tool(manager.delete_slide, parse_docstring=parse_docstring),
    tool(manager.set_slide_background, parse_docstring=parse_docstring),
    tool(manager.apply_slide_template, parse_docstring=parse_docstring),
    tool(manager.add_text_block, parse_docstring=parse_docstring),
    tool(manager.edit_text_content, parse_docstring=parse_docstring),
    tool(manager.format_text_style, parse_docstring=parse_docstring),
    tool(manager.insert_image, parse_docstring=parse_docstring),
    tool(manager.replace_image, parse_docstring=parse_docstring),
    tool(manager.create_chart, parse_docstring=parse_docstring),
    tool(manager.modify_chart_data, parse_docstring=parse_docstring),
    tool(manager.create_table, parse_docstring=parse_docstring),
    tool(manager.edit_table_cell, parse_docstring=parse_docstring),
    tool(manager.delete_shape, parse_docstring=parse_docstring),
]


def str2bool(value: str) -> bool:
    return str(value).lower() in ('true', '1', 'yes', 'y', 'on')


if str2bool(os.environ.get('USE_PROXY_URLS', 'False')):
    http_async_client = httpx.AsyncClient(proxy='http://127.0.0.1:1080')
    http_client = httpx.Client(proxy='http://127.0.0.1:1080')
else:
    http_async_client = None
    http_client = None

llm = init_chat_model(
    'gpt-4o-mini', model_provider='openai', http_client=http_client, http_async_client=http_async_client
)
llm_with_tools = llm.bind_tools(tools)


def signal_handler(sig, frame):
    print('\nClosing presentation and exiting...')
    manager.close()
    exit(0)


signal.signal(signal.SIGINT, signal.SIG_DFL)


def process_user_request(query: str, history: ChatMessageHistory):
    history.add_user_message(query)
    response = llm_with_tools.invoke(history.messages)
    history.add_ai_message(response.content)
    second_call_query = ''
    print(response.content)
    if tool_calls := response.additional_kwargs.get('tool_calls', []):
        print(f'🛠️ Executing {len(tool_calls)} tool calls')
        for tool_call_dict in tool_calls:
            tool_call = ChatCompletionOutputToolCall(
                id=tool_call_dict['id'],
                function=ChatCompletionOutputFunctionDefinition(**tool_call_dict['function']),
                type=tool_call_dict['type'],
            )
            print(f'🔧 Executing tool call: {tool_call.function}')
            function_name = tool_call.function.name
            arguments = json.loads(tool_call.function.arguments)
            if hasattr(manager, function_name):
                try:
                    if tool_call.function.name in [
                        'get_presentation_info',
                        'get_slide_details',
                        'analyze_slide_design',
                    ]:
                        second_call_query += str(getattr(manager, function_name)(**arguments))
                    else:
                        result = getattr(manager, function_name)(**arguments)
                except Exception as e:
                    print(f'Error executing {function_name}: {str(e)}')
            else:
                print(f'⚠️ Unknown function: {function_name}')
    if second_call_query:
        second_call_query += f'\nВ ЭТОМ ЗАПРОСЕ ТЫ БОЛЬШЕ НЕ ИМЕЕШЬ ПРАВА ИСПОЛЬЗОВАТЬ get_presentation_info, get_slide_details, или analyze_slide_design, потому что ты уже использовал их в предыдущем запросе. Ответь на запрос: {query}'
        process_user_request(second_call_query, history)


# Create output directory
Path('test_conversation').mkdir(exist_ok=True)

# Initialize conversation history
history = ChatMessageHistory()
message_counter = 1

# Start conversation loop
signal.signal(signal.SIGINT, signal_handler)
print("First, let's start with the info about the presentation.")
process_user_request(
    f'Это информация о презентации: {manager.get_presentation_info()}.\nПодробно опишите, чему посвящена презентация, в каком стиле она выполнена и что в ней представлено.',
    history,
)
print('Start conversation (Press Ctrl+C to exit)')
while True:
    try:
        user_input = input('\nYour request: ')
        process_user_request(user_input, history)

        # Save the presentation with numbered filename
        save_path = f'test_conversation/win_test{message_counter}.pptx'
        manager.save(save_path)
        print(f'Saved presentation as {save_path}')

        message_counter += 1
    except KeyboardInterrupt:
        print('\nClosing presentation and exiting...')
        manager.close()
        break
