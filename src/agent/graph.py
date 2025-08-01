"""PowerPoint ReAct Agent Graph.

A ReAct agent that uses LLM-generated C# code to edit PowerPoint files
using the .NET Open XML SDK.
"""

from __future__ import annotations

import asyncio
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from typing import Annotated, Literal, Optional, Sequence

import aiofiles
from langchain_core.messages import AnyMessage, ToolMessage
from langchain_core.runnables import RunnableConfig
from langchain_openai import ChatOpenAI
from langgraph.graph import StateGraph, add_messages


def preserve_value(current, update):
    """Reducer that preserves the current value if update is None."""
    return update if update is not None else current


@dataclass
class State:
    """Graph state for PowerPoint editing.

    Attributes:
        messages: Conversation history between user and agent
        pptx_file_path: Path to the PowerPoint file being edited
    """
    messages: Annotated[Sequence[AnyMessage], add_messages] = field(
        default_factory=list
    )
    pptx_file_path: Annotated[Optional[str],
                              preserve_value] = None


SYSTEM_PROMPT = """You are a PowerPoint editing assistant that uses C# code with the .NET Open XML SDK to modify presentations.

When editing PowerPoint files:
1. ALWAYS explain what you're about to do before generating code (but don't include any mention of code in your explanations)
2. Generate C# code that focuses on ONE slide at a time
3. Use clear variable names and include error handling
4. After each operation, report the results clearly
5. If an error occurs, explain it in user-friendly terms and then end.

When generating C# code:
- Assume the PresentationDocument is already open and available as 'presentation'
- The slide collection is available as 'presentation.PresentationPart.Presentation.SlideIdList'
- Use 0-based indexing for slides (slide 1 = index 0)
- Include appropriate null checks
- Focus on the specific changes requested for each slide

Example code structure:
```csharp
// For slide operations
var slideId = presentation.PresentationPart.Presentation.SlideIdList.ChildElements[slideIndex] as SlideId;
var slidePart = presentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
var slide = slidePart.Slide;

// Your modifications here
```

Always be conversational and helpful in your responses."""


async def execute_csharp_code(code: str, pptx_file_path: str) -> dict:
    """Execute C# code that modifies a PowerPoint file.

    Args:
        code: C# code to execute (will be wrapped in template)
        pptx_file_path: Path to the PowerPoint file to modify

    Returns:
        Dict with 'success' bool and 'output' or 'error' string
    """
    # Read the C# template
    template_path = os.path.join(os.path.dirname(__file__), "pptx_template.cs")
    if not await asyncio.to_thread(os.path.exists, template_path):
        # Create a basic template if it doesn't exist
        template_content = """#r "nuget: DocumentFormat.OpenXml, 3.0.0"

using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

string filePath = Args[0];

try
{
    using (PresentationDocument presentation = PresentationDocument.Open(filePath, true))
    {
        // USER_CODE_START
        {CODE}
        // USER_CODE_END
        
        Console.WriteLine("Successfully executed PowerPoint modifications");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    Environment.Exit(1);
}
"""
        async with aiofiles.open(template_path, 'w') as f:
            await f.write(template_content)
    else:
        async with aiofiles.open(template_path, 'r') as f:
            template_content = await f.read()

    # Replace the code placeholder
    full_code = template_content.replace("{CODE}", code)

    # Create a temporary C# file
    def create_temp_file():
        with tempfile.NamedTemporaryFile(mode='w', suffix='.cs', delete=False) as f:
            return f.name

    cs_file = await asyncio.to_thread(create_temp_file)

    async with aiofiles.open(cs_file, 'w') as f:
        await f.write(full_code)

    try:
        # First check if dotnet-script is installed
        check_result = await asyncio.to_thread(
            subprocess.run,
            ['dotnet', 'tool', 'list', '--global'],
            capture_output=True,
            text=True
        )

        has_dotnet_script = 'dotnet-script' in check_result.stdout

        if not has_dotnet_script:
            # Try to run directly with dotnet run instead
            # Create a simple project file
            proj_content = """<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net9.0</TargetFramework>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.0.0" />
  </ItemGroup>
</Project>"""

            # Create project in temp directory
            temp_dir = await asyncio.to_thread(tempfile.mkdtemp)
            proj_file = os.path.join(temp_dir, "temp.csproj")
            cs_file_in_dir = os.path.join(temp_dir, "Program.cs")

            # Write project file
            async with aiofiles.open(proj_file, 'w') as f:
                await f.write(proj_content)

            # Modify the code to have a proper Main method
            modified_code = full_code.replace(
                '#r "nuget: DocumentFormat.OpenXml, 3.0.0"\n\n', '')
            modified_code = modified_code.replace('string filePath = Args[0];', '''public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Error: Please provide the PowerPoint file path");
            Environment.Exit(1);
        }
        string filePath = args[0];''')
            modified_code = modified_code.rstrip() + '\n    }\n}'

            # Write the C# file
            async with aiofiles.open(cs_file_in_dir, 'w') as f:
                await f.write(modified_code)

            # First restore packages
            restore_result = await asyncio.to_thread(
                subprocess.run,
                ['dotnet', 'restore', proj_file],
                capture_output=True,
                text=True,
                timeout=30,
                cwd=temp_dir
            )

            if restore_result.returncode != 0:
                await asyncio.to_thread(shutil.rmtree, temp_dir)
                return {"success": False, "error": f"Package restore failed: {restore_result.stderr}"}

            # Build first to get better error messages
            build_result = await asyncio.to_thread(
                subprocess.run,
                ['dotnet', 'build', proj_file, '--no-restore'],
                capture_output=True,
                text=True,
                timeout=30,
                cwd=temp_dir
            )

            if build_result.returncode != 0:
                # Read the generated C# file to help debug
                async with aiofiles.open(cs_file_in_dir, 'r') as f:
                    generated_code = await f.read()
                await asyncio.to_thread(shutil.rmtree, temp_dir)
                # Get line numbers for better error reporting
                lines = generated_code.split('\n')
                numbered_lines = [f"{i+1}: {line}" for i, line in enumerate(lines)]
                code_with_lines = '\n'.join(numbered_lines[:50])  # Show first 50 lines
                
                return {"success": False, "error": f"Build failed:\n{build_result.stderr}\n{build_result.stdout}\n\nGenerated code (first 50 lines):\n{code_with_lines}"}

            # Run with dotnet run
            result = await asyncio.to_thread(
                subprocess.run,
                ['dotnet', 'run', '--project', proj_file,
                    '--no-build', '--', pptx_file_path],
                capture_output=True,
                text=True,
                timeout=60,
                cwd=temp_dir
            )

            # Clean up temp directory
            await asyncio.to_thread(shutil.rmtree, temp_dir)

        else:
            # Try to run with dotnet script
            result = await asyncio.to_thread(
                subprocess.run,
                ['dotnet', 'script', cs_file, '--', pptx_file_path],
                capture_output=True,
                text=True,
                timeout=30
            )

        if result.returncode == 0:
            return {"success": True, "output": result.stdout}
        else:
            # If dotnet script fails, provide helpful error message
            error_msg = result.stderr or result.stdout

            # Check if it's a missing package error
            if "DocumentFormat.OpenXml" in error_msg:
                return {"success": False, "error": "DocumentFormat.OpenXml package not found. Please install it:\n1. Run: dotnet tool install -g dotnet-script\n2. Create a script.csx file with: #r \"nuget: DocumentFormat.OpenXml, 3.0.0\"\n3. Run: dotnet script script.csx to install the package"}

            # Check if dotnet script is not installed
            if "dotnet-script" in error_msg or "No executable found" in error_msg:
                return {"success": False, "error": "dotnet-script not installed. Please run: dotnet tool install -g dotnet-script"}

            return {"success": False, "error": f"Execution error: {error_msg}"}

    except subprocess.TimeoutExpired:
        return {"success": False, "error": "Code execution timed out after 30 seconds"}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        # Clean up temporary files
        if await asyncio.to_thread(os.path.exists, cs_file):
            await asyncio.to_thread(os.unlink, cs_file)
        exe_file = cs_file.replace('.cs', '.exe')
        if await asyncio.to_thread(os.path.exists, exe_file):
            await asyncio.to_thread(os.unlink, exe_file)


async def pptx_tool(code: str, pptx_file_path: str) -> str:
    """Execute C# code to modify a PowerPoint presentation.

    Args:
        code: C# code that modifies the presentation (will be executed within a template)
        pptx_file_path: Path to the PowerPoint file to modify

    Returns:
        Success message or error description
    """
    result = await execute_csharp_code(code, pptx_file_path)

    if result["success"]:
        return f"Code executed successfully. Output: {result['output']}"
    else:
        return f"Execution failed: {result['error']}"


async def llm_node(state: State, config: RunnableConfig) -> dict:
    """LLM node that decides actions and generates code."""
    # Initialize the LLM with tools
    llm = ChatOpenAI(model="gpt-4.1", temperature=0)

    # Bind the tool with the current file path
    if state.pptx_file_path:
        tools = [
            {
                "type": "function",
                "function": {
                    "name": "execute_pptx_code",
                    "description": "Execute C# code to modify the PowerPoint presentation",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "code": {
                                "type": "string",
                                "description": "C# code to execute (will be run inside a template with presentation already open)"
                            }
                        },
                        "required": ["code"]
                    }
                }
            }
        ]
        llm_with_tools = llm.bind_tools(tools)
    else:
        llm_with_tools = llm

    # Add system prompt to messages
    messages = [{"role": "system", "content": SYSTEM_PROMPT}] + state.messages

    # Get LLM response
    response = await llm_with_tools.ainvoke(messages)

    return {"messages": [response]}


async def tools_node(state: State, config: RunnableConfig) -> dict:
    """Execute the tools called by the LLM."""
    # Get the last message (should be from AI with tool calls)
    last_message = state.messages[-1]

    if not hasattr(last_message, 'tool_calls') or not last_message.tool_calls:
        return {"messages": []}

    tool_messages = []

    for tool_call in last_message.tool_calls:
        if tool_call["name"] == "execute_pptx_code":
            # Execute the C# code
            result = await pptx_tool(
                code=tool_call["args"]["code"],
                pptx_file_path=state.pptx_file_path
            )

            # Create tool message with result
            tool_message = ToolMessage(
                content=result,
                tool_call_id=tool_call["id"]
            )
            tool_messages.append(tool_message)

    return {"messages": tool_messages}


def should_continue(state: State) -> Literal["tools", "end"]:
    """Determine whether to continue to tools or end."""
    messages = state.messages
    last_message = messages[-1]

    # If there are tool calls, continue to tools node
    if hasattr(last_message, 'tool_calls') and last_message.tool_calls:
        return "tools"

    # Otherwise end
    return "end"


# Build the graph
workflow = StateGraph(State)

# Add nodes
workflow.add_node("llm", llm_node)
workflow.add_node("tools", tools_node)

# Add edges
workflow.add_edge("__start__", "llm")
workflow.add_conditional_edges(
    "llm",
    should_continue,
    {
        "tools": "tools",
        "end": "__end__"
    }
)
workflow.add_edge("tools", "llm")

# Compile the graph
graph = workflow.compile(name="PowerPoint Editor")
