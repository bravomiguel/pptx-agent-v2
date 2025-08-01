"""Test script for PowerPoint ReAct agent."""

import asyncio
from langchain_core.messages import HumanMessage
from src.agent.graph import graph


async def test_agent():
    """Test the PowerPoint editing agent."""
    # Example test cases
    test_cases = [
        {
            "file_path": "test_presentation.pptx",
            "message": "Add a title 'Welcome to Our Presentation' to slide 1"
        },
        {
            "file_path": "test_presentation.pptx", 
            "message": "Change the background color of slide 2 to light blue"
        },
        {
            "file_path": "test_presentation.pptx",
            "message": "Add a new slide with title 'Thank You' at the end"
        }
    ]
    
    print("PowerPoint ReAct Agent Test\n" + "="*50)
    
    for i, test in enumerate(test_cases, 1):
        print(f"\nTest {i}: {test['message']}")
        print("-" * 50)
        
        # Initialize state
        initial_state = {
            "messages": [HumanMessage(content=test["message"])],
            "pptx_file_path": test["file_path"]
        }
        
        try:
            # Run the agent
            result = await graph.ainvoke(initial_state)
            
            # Print the conversation
            for msg in result["messages"]:
                role = msg.__class__.__name__
                print(f"\n{role}: {msg.content}")
                
        except Exception as e:
            print(f"Error: {e}")
    
    print("\n" + "="*50)
    print("Testing complete!")


if __name__ == "__main__":
    # Note: This requires a test PowerPoint file to exist
    print("Note: This test requires:")
    print("1. .NET SDK installed (for dotnet script)")
    print("2. A test PowerPoint file named 'test_presentation.pptx'")
    print("3. OpenAI API key set in environment")
    print("\nPress Enter to continue or Ctrl+C to cancel...")
    input()
    
    asyncio.run(test_agent())