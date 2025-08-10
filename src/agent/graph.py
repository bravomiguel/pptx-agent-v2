"""PowerPoint ReAct Agent Graph.

A ReAct agent that uses LLM-generated C# code to edit PowerPoint files
using the .NET Open XML SDK.
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from typing import Annotated, Literal, Optional, Sequence, List

import aiofiles
from langchain_core.messages import AnyMessage, ToolMessage
from langchain_core.runnables import RunnableConfig
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI
from langgraph.graph import StateGraph, add_messages
from langgraph.prebuilt import InjectedState, ToolNode
from typing_extensions import Annotated as TypeAnnotated


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

## Your Available Tools:
1. **read_pptx_structure**: Get overview of entire presentation (slide count, titles, layouts)
2. **read_slide_details**: Get detailed content from specific slides (text, formatting, positions, semantic anchors)
3. **execute_pptx_code**: Run C# code to modify the presentation

## CRITICAL WORKFLOW RULES:
**ALWAYS read before editing when the task involves:**
- Copying formatting between slides
- Modifying existing content (bullets, text, etc.)
- Finding specific text or elements
- Making slides consistent
- Any operation that references existing content

## Reading Tool Usage:
- Use **read_pptx_structure** first to understand the presentation and locate relevant slides
- Use **read_slide_details** to get specific content, formatting, and element structure
- Read multiple slides in one call when comparing: read_slide_details([2, 3, 4])

## Memory and Re-reading:
- **IMPORTANT**: Check your conversation history before reading - if you've already read a slide, use that information
- Look for the JSON response in your previous tool calls
- Only re-read a slide if:
  - You've modified it since last reading (anchors will have changed)
  - You need information not captured in the previous read
  - The user explicitly asks for fresh information
- After modifying a slide, you MUST re-read it if you need updated anchors for further edits

## Understanding the JSON Response:
The reading tools return JSON with this structure:
```json
{
  "SlideNumber": 1,
  "Elements": [
    {
      "Anchor": {
        "Anchor": "slide1_title0_abc123",  // Unique identifier
        "Type": "title",                   // Element type
        "Path": "slide[1].title[0]"        // Structural path
      },
      "Content": "Text content",
      "Children": [],  // For bullets, contains sub-items
      "Formatting": {}, // Font, size, colors, etc.
      "Position": {}    // X, Y, width, height
    }
  ]
}
```

## Semantic Anchors:
- Format: `slide{num}_{type}{index}_{hash}`
- Anchors are unique identifiers for each element
- They change when content is modified (this is expected)
- Use anchors to identify elements, but remember they update after edits

## Common Task Patterns:

**"Make slide X bullets more concise":**
1. read_slide_details([X]) to see current bullets
2. Analyze the content structure and Children arrays
3. Generate code to modify each bullet by index
4. Explain which bullets you're changing

**"Format slide X like slide Y":**
1. read_slide_details([X, Y]) to get both slides
2. Extract formatting from slide Y's JSON
3. Apply formatting to slide X elements
4. Verify the changes if needed

**"Find and modify specific text":**
1. read_pptx_structure() to locate slides containing the text
2. read_slide_details() for those specific slides
3. Use element indices or content matching to target the right element
4. Make the modification

## When working with PowerPoint files:
1. ALWAYS explain what you're about to do before generating code (but don't include any mention of code in your explanations)
2. For content-aware edits, FIRST read the relevant slides
3. Generate C# code that focuses on ONE slide at a time
4. Use clear variable names and include error handling
5. After each operation, report the results clearly
6. If an error occurs, try to fix once only, and if it fails again, provide explanation and then end, no more retries

## When generating C# code:
- Your code will be injected inside a Main method with the PresentationDocument already open as 'presentation'
- DO NOT include 'using' statements - all necessary namespaces are already imported
- DO NOT use 'return' statements - use Console.WriteLine() to output results
- Available direct imports: System, System.Linq, DocumentFormat.OpenXml.Packaging, DocumentFormat.OpenXml.Presentation, DocumentFormat.OpenXml
- Namespace aliases available: 
  - P = DocumentFormat.OpenXml.Presentation (use P.Shape, P.SlideId, etc.)
  - D = DocumentFormat.OpenXml.Drawing (use D.Paragraph, D.Run, D.Text, etc.)
  - **IMPORTANT**: Use D for Drawing types, NOT A (which doesn't exist)
- Common types and their namespaces:
  - Shape, SlideId, SlidePart: Use with P prefix (P.Shape) or directly (Shape is also imported)
  - Paragraph, Run, Text: ALWAYS use with D prefix (D.Paragraph, D.Run, D.Text)
- The slide collection is available as 'presentation.PresentationPart.Presentation.SlideIdList'
- Use 0-based indexing for slides (slide 1 = index 0)
- Include appropriate null checks

## Element Targeting:
- Use indices when you know the exact position (from reading)
- Use content matching when searching for specific text
- For bullets, remember to check the Children array for sub-bullets
- When multiple elements match, ask the user for clarification

Example code structure:
```csharp
// For slide operations
var slideId = presentation.PresentationPart.Presentation.SlideIdList.ChildElements[slideIndex] as SlideId;
var slidePart = presentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
var slide = slidePart.Slide;

// Example: Modifying text in a shape
var shape = slide.Descendants<P.Shape>().FirstOrDefault();  // P.Shape or just Shape
if (shape?.TextBody != null)
{
    var paragraph = shape.TextBody.GetFirstChild<D.Paragraph>();  // MUST use D. for Drawing types
    if (paragraph != null)
    {
        paragraph.RemoveAllChildren<D.Run>();  // D.Run, NOT A.Run
        var run = new D.Run(new D.Text("New text"));  // D.Run and D.Text
        paragraph.Append(run);
    }
}

Console.WriteLine("Operation completed successfully");
```

## Common PowerPoint Structure Rules
To avoid validation errors, remember these structural requirements:

1. **Shape Tree (spTree) Requirements**: Every spTree element must contain its required child elements, including nvGrpSpPr (non-visual group shape properties). Never remove these required elements when modifying slide content.

Always be conversational and helpful in your responses."""


async def execute_csharp_code(code: str, pptx_file_path: str) -> dict:
    """Execute C# code that modifies a PowerPoint file.

    Args:
        code: C# code to execute (will be wrapped in template)
        pptx_file_path: Path to the PowerPoint file to modify

    Returns:
        Dict with 'success' bool and 'output' or 'error' string
    """
    # Get the path to the .NET project
    project_dir = os.path.join(os.path.dirname(__file__), "PptxEditor")
    program_file = os.path.join(project_dir, "Program.cs")
    
    # Read the C# template
    async with aiofiles.open(program_file, 'r') as f:
        template_content = await f.read()
    
    # Replace the code placeholder
    full_code = template_content.replace("// {CODE}", code)
    
    # Create a temporary directory for execution
    temp_dir = await asyncio.to_thread(tempfile.mkdtemp)
    temp_program = os.path.join(temp_dir, "Program.cs")
    temp_project = os.path.join(temp_dir, "PptxEditor.csproj")
    
    # Copy the project file
    project_file = os.path.join(project_dir, "PptxEditor.csproj")
    await asyncio.to_thread(shutil.copy2, project_file, temp_project)
    
    # Write the modified Program.cs
    async with aiofiles.open(temp_program, 'w') as f:
        await f.write(full_code)

    try:
        # First restore packages
        restore_result = await asyncio.to_thread(
            subprocess.run,
            ['dotnet', 'restore', temp_project],
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
            ['dotnet', 'build', temp_project, '--no-restore'],
            capture_output=True,
            text=True,
            timeout=30,
            cwd=temp_dir
        )

        if build_result.returncode != 0:
            # Read the generated C# file to help debug
            async with aiofiles.open(temp_program, 'r') as f:
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
            ['dotnet', 'run', '--project', temp_project,
                '--no-build', '--', pptx_file_path],
            capture_output=True,
            text=True,
            timeout=60,
            cwd=temp_dir
        )

        # Clean up temp directory
        await asyncio.to_thread(shutil.rmtree, temp_dir)

        if result.returncode == 0:
            return {"success": True, "output": result.stdout}
        elif result.returncode == 2:
            # Validation error - parse the validation messages
            validation_errors = result.stdout
            return {"success": False, "error": f"Validation failed - the modifications would corrupt the PowerPoint file:\n{validation_errors}"}
        else:
            error_msg = result.stderr or result.stdout
            return {"success": False, "error": f"Execution error: {error_msg}"}

    except subprocess.TimeoutExpired:
        await asyncio.to_thread(shutil.rmtree, temp_dir)
        return {"success": False, "error": "Code execution timed out after 60 seconds"}
    except Exception as e:
        if await asyncio.to_thread(os.path.exists, temp_dir):
            await asyncio.to_thread(shutil.rmtree, temp_dir)
        return {"success": False, "error": str(e)}


@tool
async def execute_pptx_code(
    code: str,
    state: TypeAnnotated[State, InjectedState]
) -> str:
    """Execute C# code to modify the PowerPoint presentation.
    
    Args:
        code: C# code to execute (will be run inside a template with presentation already open)
        state: The graph state (injected automatically)
    
    Returns:
        Success message or error description
    """
    pptx_file_path = state.pptx_file_path
    if not pptx_file_path:
        return "Error: No PowerPoint file path provided in state"
    
    result = await execute_csharp_code(code, pptx_file_path)

    if result["success"]:
        return f"Code executed successfully. Output: {result['output']}"
    else:
        return f"Execution failed: {result['error']}"


async def execute_reading_code(code: str, pptx_file_path: str) -> dict:
    """Execute C# code that reads from a PowerPoint file.
    
    Args:
        code: C# code to execute (will be wrapped in template)
        pptx_file_path: Path to the PowerPoint file to read
    
    Returns:
        Dict with 'success' bool and 'output' or 'error' string
    """
    # Get the path to the .NET project
    project_dir = os.path.join(os.path.dirname(__file__), "PptxEditor")
    
    # Create a temporary directory for execution
    temp_dir = await asyncio.to_thread(tempfile.mkdtemp)
    temp_program = os.path.join(temp_dir, "ReadProgram.cs")
    temp_project = os.path.join(temp_dir, "PptxEditor.csproj")
    
    # Copy the project file and PptxReader.cs
    project_file = os.path.join(project_dir, "PptxEditor.csproj")
    reader_file = os.path.join(project_dir, "PptxReader.cs")
    await asyncio.to_thread(shutil.copy2, project_file, temp_project)
    await asyncio.to_thread(shutil.copy2, reader_file, os.path.join(temp_dir, "PptxReader.cs"))
    
    # Create a reading program
    read_program = f"""
using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

public class ReadProgram
{{
    public static void Main(string[] args)
    {{
        if (args.Length < 1)
        {{
            Console.WriteLine("Error: Please provide the PowerPoint file path");
            Environment.Exit(1);
        }}
        
        string filePath = args[0];
        
        try
        {{
            {code}
        }}
        catch (Exception ex)
        {{
            Console.WriteLine($"Error: {{ex.Message}}");
            Environment.Exit(1);
        }}
    }}
}}
"""
    
    # Write the reading program
    async with aiofiles.open(temp_program, 'w') as f:
        await f.write(read_program)
    
    try:
        # Build the project
        build_result = await asyncio.to_thread(
            subprocess.run,
            ['dotnet', 'build', temp_project, '/p:StartupObject=ReadProgram'],
            capture_output=True,
            text=True,
            timeout=30,
            cwd=temp_dir
        )
        
        if build_result.returncode != 0:
            await asyncio.to_thread(shutil.rmtree, temp_dir)
            return {"success": False, "error": f"Build failed: {build_result.stderr}"}
        
        # Run the reading program
        result = await asyncio.to_thread(
            subprocess.run,
            ['dotnet', 'run', '--project', temp_project,
                '--no-build', '--', pptx_file_path],
            capture_output=True,
            text=True,
            timeout=60,
            cwd=temp_dir
        )
        
        # Clean up temp directory
        await asyncio.to_thread(shutil.rmtree, temp_dir)
        
        if result.returncode == 0:
            return {"success": True, "output": result.stdout}
        else:
            error_msg = result.stderr or result.stdout
            return {"success": False, "error": f"Execution error: {error_msg}"}
    
    except subprocess.TimeoutExpired:
        await asyncio.to_thread(shutil.rmtree, temp_dir)
        return {"success": False, "error": "Code execution timed out"}
    except Exception as e:
        if await asyncio.to_thread(os.path.exists, temp_dir):
            await asyncio.to_thread(shutil.rmtree, temp_dir)
        return {"success": False, "error": str(e)}


@tool
async def read_pptx_structure(
    state: TypeAnnotated[State, InjectedState]
) -> str:
    """Read the overall structure of the PowerPoint presentation.
    
    Returns a JSON overview including:
    - Total number of slides
    - Slide titles and layouts
    - High-level element information with semantic anchors
    
    Args:
        state: The graph state (injected automatically)
    
    Returns:
        JSON string with presentation structure
    """
    pptx_file_path = state.pptx_file_path
    if not pptx_file_path:
        return "Error: No PowerPoint file path provided in state"
    
    code = f"""
string result = PptxReader.ReadStructure(filePath);
Console.WriteLine(result);
"""
    
    result = await execute_reading_code(code, pptx_file_path)
    
    if result["success"]:
        return result['output']
    else:
        return f"Failed to read presentation structure: {result['error']}"


@tool
async def read_slide_details(
    slide_numbers: List[int],
    state: TypeAnnotated[State, InjectedState]
) -> str:
    """Read detailed content from specific slides.
    
    Returns detailed information including:
    - Full text content with hierarchy
    - Formatting details
    - Element positions
    - Semantic anchors for precise targeting
    
    Args:
        slide_numbers: List of slide numbers to read (1-based indexing)
        state: The graph state (injected automatically)
    
    Returns:
        JSON string with detailed slide information
    """
    pptx_file_path = state.pptx_file_path
    if not pptx_file_path:
        return "Error: No PowerPoint file path provided in state"
    
    # Convert to C# array
    numbers_str = ", ".join(str(n) for n in slide_numbers)
    code = f"""
int[] slideNumbers = new int[] {{ {numbers_str} }};
string result = PptxReader.ReadSlideDetails(filePath, slideNumbers);
Console.WriteLine(result);
"""
    
    result = await execute_reading_code(code, pptx_file_path)
    
    if result["success"]:
        return result['output']
    else:
        return f"Failed to read slide details: {result['error']}"


async def llm_node(state: State, config: RunnableConfig) -> dict:
    """LLM node that decides actions and generates code."""
    # Initialize the LLM with tools
    # llm = ChatOpenAI(model="gpt-5")
    llm = ChatOpenAI(model="gpt-4.1", temperature=0)

    # Bind the tools with the current file path
    if state.pptx_file_path:
        llm_with_tools = llm.bind_tools([execute_pptx_code, read_pptx_structure, read_slide_details])
    else:
        llm_with_tools = llm

    # Add system prompt to messages
    messages = [{"role": "system", "content": SYSTEM_PROMPT}] + state.messages

    # Get LLM response
    response = await llm_with_tools.ainvoke(messages)

    return {"messages": [response]}


# Create the ToolNode with our decorated tools
tools_node = ToolNode([execute_pptx_code, read_pptx_structure, read_slide_details])


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
