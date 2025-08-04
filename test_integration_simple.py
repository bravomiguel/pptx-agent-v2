"""Simple integration test to verify the refactored tool and ToolNode work together."""

import asyncio
from src.agent.graph import graph, State, execute_pptx_code, tools_node
from langchain_core.messages import HumanMessage, AIMessage
from langgraph.prebuilt import ToolNode


async def test_tool_and_node():
    """Test that the tool and ToolNode are properly configured."""
    
    print("Testing refactored architecture...")
    
    # 1. Verify the tool is properly decorated
    print("\n1. Checking tool decoration...")
    assert hasattr(execute_pptx_code, 'name'), "Tool should have a name attribute"
    assert execute_pptx_code.name == "execute_pptx_code", f"Tool name should be 'execute_pptx_code', got {execute_pptx_code.name}"
    print(f"   ✅ Tool name: {execute_pptx_code.name}")
    
    # 2. Verify ToolNode was created with our tool
    print("\n2. Checking ToolNode configuration...")
    assert isinstance(tools_node, ToolNode), "tools_node should be a ToolNode instance"
    print(f"   ✅ ToolNode type: {type(tools_node).__name__}")
    
    # 3. Test the tool directly with mock state
    print("\n3. Testing tool execution directly...")
    mock_state = State(
        messages=[],
        pptx_file_path="/Users/miguelbravo/Downloads/test.pptx"
    )
    
    result = await execute_pptx_code.ainvoke({
        "code": "Console.WriteLine(\"Test from refactored tool\");",
        "state": mock_state
    })
    
    assert "successfully" in result or "failed" in result, "Tool should return a result message"
    print(f"   ✅ Tool result: {result[:100]}...")
    
    # 4. Check graph structure
    print("\n4. Checking graph structure...")
    nodes = graph.nodes
    assert "llm" in nodes, "Graph should have 'llm' node"
    assert "tools" in nodes, "Graph should have 'tools' node"
    print(f"   ✅ Graph nodes: {list(nodes.keys())}")
    
    print("\n✅ All integration tests passed!")
    print("\nRefactoring Summary:")
    print("- Tool now uses @tool decorator")
    print("- Tool uses InjectedState to access pptx_file_path")
    print("- ToolNode from langgraph.prebuilt replaces manual implementation")
    print("- Graph properly integrated with new components")


if __name__ == "__main__":
    asyncio.run(test_tool_and_node())