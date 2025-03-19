import getpass
import os
import json
import signal
from pathlib import Path
from dotenv import load_dotenv
from langchain.chat_models import init_chat_model
from langchain_core.tools import tool
from langchain_community.chat_message_histories import ChatMessageHistory
from huggingface_hub.inference._generated.types.chat_completion import ChatCompletionOutputToolCall, ChatCompletionOutputFunctionDefinition
from tools.pptx_tools import PPTXManager


load_dotenv()

if not os.environ.get("OPENAI_API_KEY"):
    os.environ["OPENAI_API_KEY"] = getpass.getpass("Enter API key for OpenAI: ")

manager = PPTXManager("test_sources/test_dit.pptx")
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
    tool(manager.delete_shape, parse_docstring=parse_docstring)
]

llm = init_chat_model("gpt-4o-mini", model_provider="openai")

llm_with_tools = llm.bind_tools(tools)

def signal_handler(sig, frame):
    print("\nClosing presentation and exiting...")
    manager.close()
    exit(0)

signal.signal(signal.SIGINT, signal.SIG_DFL)

def process_user_request(query: str, history: ChatMessageHistory):
    history.add_user_message(query)
    response = llm_with_tools.invoke(history.messages)
    history.add_ai_message(response.content)
    
    print(response.content)
    if tool_calls := response.additional_kwargs.get("tool_calls", []):
        print(f"🛠️ Executing {len(tool_calls)} tool calls")
        for tool_call_dict in tool_calls:
            tool_call = ChatCompletionOutputToolCall(
                        id=tool_call_dict["id"],
                        function=ChatCompletionOutputFunctionDefinition(**tool_call_dict["function"]),
                        type=tool_call_dict["type"]
                    )
            print(f"🔧 Executing tool call: {tool_call.function}")
            function_name = tool_call.function.name
            arguments = json.loads(tool_call.function.arguments)

            if hasattr(manager, function_name):
                try:
                    if tool_call.function.name in ["get_presentation_info", "get_slide_details", "analyze_slide_design"]:
                        process_user_request(getattr(manager, function_name)(**arguments), history)
                    else:
                        result = getattr(manager, function_name)(**arguments)
                except Exception as e:
                    print(f"Error executing {function_name}: {str(e)}")
            else:
                print(f"⚠️ Unknown function: {function_name}")

# Create output directory
Path("test_conversation").mkdir(exist_ok=True)

# Initialize conversation history
history = ChatMessageHistory()
message_counter = 1

# Start conversation loop
signal.signal(signal.SIGINT, signal_handler)
print("First, let's start with the info about the presentation.")
process_user_request(f"This is the info about presentation: {manager.get_presentation_info()}. Describe in detail what the presentation is about, what styles it is in, and what it presents.", history)
print("Start conversation (Press Ctrl+C to exit)")
while True:
    try:
        user_input = input("\nYour request: ")
        process_user_request(user_input, history)
        
        # Save the presentation with numbered filename
        save_path = f"test_conversation/win_test{message_counter}.pptx"
        manager.save(save_path)
        print(f"Saved presentation as {save_path}")
        
        message_counter += 1
    except KeyboardInterrupt:
        print("\nClosing presentation and exiting...")
        manager.close()
        break