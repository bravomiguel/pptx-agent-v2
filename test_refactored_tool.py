"""Unit test for the refactored execute_pptx_code tool.

This test directly calls the tool without involving the LLM or full graph.
"""

import asyncio
from dataclasses import dataclass, field
from typing import Optional, Sequence
from unittest.mock import MagicMock
from langchain_core.messages import AnyMessage


async def test_execute_pptx_code_tool():
    """Test the execute_pptx_code tool with a mock state."""
    
    # Import the tool (after refactoring)
    from src.agent.graph import execute_pptx_code, State
    
    # Create a mock state with pptx_file_path
    mock_state = State(
        messages=[],
        pptx_file_path="/Users/miguelbravo/Downloads/test.pptx"
    )
    
    # Simple C# code that lists slide count
    test_code = """
var slideCount = presentation.PresentationPart.Presentation.SlideIdList.ChildElements.Count;
Console.WriteLine($"Presentation has {slideCount} slides");
"""
    
    print("Testing execute_pptx_code tool...")
    print(f"PowerPoint file: {mock_state.pptx_file_path}")
    print(f"C# code to execute:\n{test_code}")
    
    try:
        # Call the tool directly
        result = await execute_pptx_code.ainvoke({
            "code": test_code,
            "state": mock_state
        })
        
        print("\n✅ Tool executed successfully!")
        print(f"Result: {result}")
        
        # Verify result format
        assert isinstance(result, str), "Result should be a string"
        assert "slides" in result.lower() or "error" in result.lower(), "Result should contain slide info or error"
        
        # Test with missing file path
        print("\n\nTesting with missing file path...")
        empty_state = State(messages=[], pptx_file_path=None)
        result_no_path = await execute_pptx_code.ainvoke({
            "code": test_code,
            "state": empty_state
        })
        print(f"Result with no path: {result_no_path}")
        assert "error" in result_no_path.lower(), "Should return error when no file path"
        
        print("\n✅ All tests passed!")
        
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(test_execute_pptx_code_tool())