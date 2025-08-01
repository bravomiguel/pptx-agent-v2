# PowerPoint ReAct Agent Setup Guide

## Prerequisites

1. **Python 3.9+** installed
2. **.NET SDK** (for C# code execution)
3. **OpenAI API Key**

## Installation Steps

### 1. Install .NET SDK

#### macOS:
```bash
brew install --cask dotnet-sdk
```

#### Windows:
Download from: https://dotnet.microsoft.com/download

#### Linux:
```bash
# Ubuntu/Debian
wget https://packages.microsoft.com/config/ubuntu/20.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y dotnet-sdk-8.0
```

### 2. Install DocumentFormat.OpenXml

```bash
dotnet tool install -g dotnet-script
dotnet add package DocumentFormat.OpenXml
```

Or create a global tool manifest:
```bash
dotnet new tool-manifest
dotnet tool install dotnet-script
```

### 3. Install Python Dependencies

```bash
pip install -e .
```

### 4. Set OpenAI API Key

```bash
export OPENAI_API_KEY="your-api-key-here"
```

Or create a `.env` file:
```
OPENAI_API_KEY=your-api-key-here
```

## Usage

### Running the Agent

```python
import asyncio
from langchain_core.messages import HumanMessage
from src.agent.graph import graph

async def edit_presentation():
    state = {
        "messages": [HumanMessage(content="Add a title to slide 1")],
        "pptx_file_path": "my_presentation.pptx"
    }
    
    result = await graph.ainvoke(state)
    
    # Print conversation
    for msg in result["messages"]:
        print(f"{msg.__class__.__name__}: {msg.content}")

asyncio.run(edit_presentation())
```

### Example Commands

1. **Add content to slides:**
   - "Add a title 'Q4 Results' to slide 1"
   - "Insert a subtitle 'Financial Overview' on slide 2"

2. **Modify slide appearance:**
   - "Change the background color of slide 3 to blue"
   - "Make the title on slide 1 bold and red"

3. **Create new slides:**
   - "Add a new slide with title 'Conclusion' at the end"
   - "Insert a blank slide after slide 2"

4. **Work with slide elements:**
   - "Add a text box with 'Important Note' on slide 4"
   - "Update the bullet points on slide 5"

## Troubleshooting

### Common Issues

1. **"dotnet: command not found"**
   - Ensure .NET SDK is installed and in PATH
   - Restart terminal after installation

2. **"Could not find DocumentFormat.OpenXml.dll"**
   - Install the NuGet package globally:
     ```bash
     dotnet add package DocumentFormat.OpenXml --version 3.0.0
     ```

3. **Compilation errors**
   - Check that C# code syntax is valid
   - Ensure all required namespaces are imported

4. **File access errors**
   - Ensure the PPTX file exists and is not open in PowerPoint
   - Check file permissions

## Architecture Overview

The agent uses a ReAct (Reasoning and Acting) pattern:

1. **User Request** → Analyzed by LLM
2. **LLM generates C# code** → Focused on one slide at a time
3. **Code execution** → Via .NET subprocess
4. **Result feedback** → LLM reports success/failure
5. **Cycle continues** → Until task is complete

The system prioritizes:
- Clear user feedback at each step
- Safe file handling (preserves originals)
- Error recovery through LLM retry
- Slide-scoped operations for clarity