# PowerPoint ReAct Agent Design

## Overview
A ReAct (Reasoning and Acting) agent that uses LLM-generated C# code with the .NET Open XML SDK to edit PowerPoint files.

## Architecture

### Graph Structure
- **Pattern**: ReAct with cycles between reasoning and acting
- **Nodes**:
  - `llm_node`: Analyzes request, generates C# code for slide edits
  - `tools_node`: Executes the generated .NET Open XML SDK code
  - Flow cycles between these nodes until task is complete

### State Design
```python
@dataclass
class State:
    messages: List[Message]  # Conversation history
    pptx_file_path: str     # Path to PowerPoint file being edited
```

### Code Generation Strategy
- **Scope**: Slide-level operations (all edits for one slide per generation)
- **Language**: C# using DocumentFormat.OpenXml
- **Template Wrapper**: Generated code is injected into a pre-defined template with necessary imports and error handling

Example template:
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
// other imports...

public class PptxEditor {
    public static void Execute(string filePath) {
        using (var doc = PresentationDocument.Open(filePath, true)) {
            // {GENERATED_CODE_HERE}
        }
    }
}
```

### Execution Method
- **Approach**: Subprocess execution using `dotnet script` or compile-and-run
- **Error Handling**: 
  - Capture compilation errors → feed back to LLM for correction
  - Capture runtime exceptions → provide context for retry
- **File Safety**: Work on copies, preserve original files

### User Feedback Mechanism
- **Method**: System prompt instructs LLM to narrate actions
- **Implementation**: Natural conversation flow with status updates
- **No streaming complexity**: Just standard AI messages in conversation

Example prompt instruction:
```
As you work:
1. Always explain what you're about to do before doing it
2. Provide clear status updates after each operation
3. Report any errors in user-friendly terms
```

### Execution Flow Example
```
User: "Add titles to slides 1 and 2"
AI: "I'll add titles to slides 1 and 2. Starting with slide 1..."
→ [Generates C# code for slide 1 edits]
→ [Executes code via tools_node]
AI: "✓ Slide 1 complete! Now working on slide 2..."
→ [Generates C# code for slide 2 edits]
→ [Executes code via tools_node]
AI: "✓ Done! Both slides now have titles."
```

## Key Benefits

1. **Simplicity**: Minimal state, straightforward flow
2. **Natural Boundaries**: Slide-scoped operations match user mental model
3. **Error Isolation**: Failures on one slide don't affect completed slides
4. **Clear Progress**: Users see real-time updates through conversation
5. **Full Power**: Complete access to Open XML SDK capabilities

## Implementation Principles

- Start with minimal viable functionality
- Let the LLM handle planning through natural conversation
- No explicit action plans in state - conversation history serves this purpose
- Focus on robust error handling and clear user feedback
- Maintain file safety through defensive practices

## Future Enhancements (Post-MVP)

- Batch operations for performance
- Checkpoint system for long operations
- Template library for common operations
- Multi-file support
- Undo/redo capability