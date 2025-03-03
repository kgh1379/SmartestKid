#!/usr/bin/env python3
import os
import json
import time
from openai import OpenAI
from tools.excel_tools import ExcelTools
from tools.image_tools import ImageTools
from tools.directory_tools import DirectoryTools
from tools.word_tools import WordTools
from dotenv import load_dotenv

class ChatCompletionAssistant:
    def __init__(self, api_key=None):
        # Initialize OpenAI client with the provided API key
        self.client = OpenAI(api_key=api_key)
        
        # Initialize conversation with system message
        self.conversation = [
            {
                "role": "system",
                "content": (
                    "You are an ultra smart modeling expert. You help people build and optimize models of all sorts."
                    "You have access to various tools such as analyzing files, processing Excel and Word documents, listing directories, and running calculations."
                    "When you use a tool, always explain what you found or what happened immediately after using it. "
                    "Also, always, before using a tool, explain succinctly what tools you will use before you use them!"
                    "Be interactive and conversational - if you need to use multiple tools, discuss the results of each one before moving to the next. "
                    "In certain scenarios, you will be asked to plug numbers from one excel into another excel calculator and record. This may involve many loops writing in file A, file B, file A, etc. "
                )
            }
        ]
        self.functions = [
            {
                "type": "function",
                "function": {
                    "name": "analyze_file",
                    "description": "Analyze either an image or PDF file, using computer vision, to return valuable data or answer questions about it. Make sure you know the exact name of the file (i.e. might have to list directory tool first) before you open it!",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "question": {
                                "type": "string",
                                "description": "Question or instruction about what to analyze in the file. Be comprehensive."
                            },
                            "file_name": {
                                "type": "string",
                                "description": "Name of the file in the datalake directory."
                            }
                        },
                        "required": ["file_name","question"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "process_excel",
                    "description": "This function lets you create new excel files and edit existing ones. The input is the excel file name (file_path), as well as a set of tuples that represent data to put in the file (write_data). If the file_path doesnt exist in our data lake, it creates a new one. If the set of tuples is empty, we are not writing anything, just reading. At the end of the call, this function returns returns the post-edited state of the excel file. For example, if you just want to read the file, you can pass an empty list for write_data. However, if you're entering a value or values into an excel calculator, you dont need to do a distinct read because after you write, the function will return the new updated calculated state. write_data should be of the format {'A1': 42, 'B2': 'hello', 'C3': 'apple'}. Cell addresses MUST be in A1 notation (A1, B2, etc).",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Name or path of the Excel file to create or edit."
                            },
                            "write_data": {
                                "type": "object",
                                "description": "REQUIRED for writing. Simple dictionary mapping cell addresses to values, e.g., {'A1': 42, 'B2': 'hello', 'C3': 'apple'}. Cell addresses MUST be in A1 notation (A1, B2, etc).",
                                "additionalProperties": {
                                    "type": ["string", "number"]
                                }
                            }
                        },
                        "required": ["file_path", "write_data"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "process_excel",
                    "description": "Process an Excel file. When writing data, you MUST provide both file_path and write_data. Creates a new file if it doesn't exist.  Cell addresses MUST be in A1 notation (A1, B2, etc).",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Name or path of the Excel file to process."
                            },
                            "write_data": {
                                "type": "object",
                                "description": "REQUIRED for writing. Simple dictionary mapping cell addresses to values, e.g., {'A1': 42, 'B2': 'hello', 'C3': 'apple'}. Cell addresses MUST be in A1 notation (A1, B2, etc).",
                                "additionalProperties": {
                                    "type": ["string", "number"]
                                }
                            }
                        },
                        "required": ["file_path", "write_data"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "process_word",
                    "description": "Process a Word document - read current content and optionally write new content",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Name of the Word document (e.g., 'mydoc.docx')"
                            },
                            "content": {
                                "type": "string",
                                "description": "Optional content to write to the document"
                            }
                        },
                        "required": ["file_path"]
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "list_directory",
                    "description": "List all files and directories in the DataLake directory",
                    "parameters": {
                        "type": "object",
                        "properties": {},
                        "required": []
                    }
                }
            },
        ]
        # Initialize local tool instances.
        self.excel_tools = ExcelTools()
        self.image_tools = ImageTools()
        self.directory_tools = DirectoryTools()
        self.word_tools = WordTools()

    def send_message(self, message: str, ui_callback=None):

        loop_counter = 0

        start_time = time.time()

        if message:
            self.conversation.append({"role": "user", "content": message})
        complete_response = ""
        
        while True:  # Keep going until we get a non-tool finish_reason

            loop_counter += 1

            try:
                completion = self.client.chat.completions.create(
                    model="gpt-4o",
                    messages=self.conversation,
                    stream=True,
                    tools=self.functions,
                    tool_choice="auto"
                )
                
                final_tool_calls = {}  # Track complete tool calls
                
                for chunk in completion:
                    delta = chunk.choices[0].delta
                    
                    # Handle regular content
                    if delta.content:
                        complete_response += delta.content
                        if ui_callback:
                            ui_callback(delta.content)
                    
                    # Handle tool calls
                    if delta.tool_calls:
                        for tool_call in delta.tool_calls:
                            # Initialize if new tool call
                            if tool_call.index not in final_tool_calls:
                                final_tool_calls[tool_call.index] = {
                                    "id": tool_call.id,
                                    "type": "function",
                                    "function": {
                                        "name": tool_call.function.name,
                                        "arguments": ""
                                    }
                                }
                            
                            # Accumulate arguments
                            if tool_call.function and tool_call.function.arguments:
                                final_tool_calls[tool_call.index]["function"]["arguments"] += tool_call.function.arguments
                    
                    # Check finish reason
                    if chunk.choices[0].finish_reason == "tool_calls":
                        for tool_call in final_tool_calls.values():
                            
                            try:
                                if ui_callback:
                                    ui_callback(f"\n\nRunning {tool_call['function']['name']}...\n\n")

                                func_args = json.loads(tool_call["function"]["arguments"])
                                result = self.handle_function_call(tool_call["function"]["name"], func_args)
                                
                                # Add to conversation history
                                self.conversation.append({
                                    "role": "assistant",
                                    "content": None,
                                    "tool_calls": [tool_call]
                                })
                                self.conversation.append({
                                    "role": "tool",
                                    "tool_call_id": tool_call["id"],
                                    "content": str(result)
                                })
                            except json.JSONDecodeError as e:
                                print(f"[DEBUG] Error parsing arguments: {e}")
                        break
                    
                    elif chunk.choices[0].finish_reason:
                        print(f"[DEBUG] Finish reason: {chunk.choices[0].finish_reason}")
                        if complete_response:
                            self.conversation.append({"role": "assistant", "content": complete_response})
                        if ui_callback:
                            ui_callback({"end_of_message": True})
                        return complete_response
                
            except Exception as e:
                print(f"Error in send_message: {str(e)}")
                return f"Error: {str(e)}"

    def handle_function_call(self, func_name: str, arguments: dict) -> str:
        print(f"[DEBUG] Handling function call: {func_name} with arguments {arguments}")
        if func_name == "process_excel":
            result = self.excel_tools.process_excel(arguments.get("file_path"), arguments.get("write_data"))
        elif func_name == "process_word":
            result = self.word_tools.process_word(arguments.get("file_path"), arguments.get("content"))
        elif func_name == "analyze_file":
            result = self.image_tools.analyze_file(arguments.get("question"), arguments.get("file_name"))
        elif func_name == "list_directory":
            result = self.directory_tools.list_directory()
        else:
            result = f"Function {func_name} not implemented."
        return str(result)

    def cleanup(self):
        self.excel_tools.cleanup()
        self.word_tools.cleanup()
