"""Test script for the refactored dotnet implementation."""

import asyncio
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from agent.graph import execute_csharp_code

async def test_pptx_editor():
    """Test the PowerPoint editor with hardcoded C# code."""
    
    # Simple C# code that lists slides in the presentation
    test_code = """
// List all slides in the presentation
var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
Console.WriteLine($"Total slides in presentation: {slideIdList.Count()}");

int slideIndex = 0;
foreach (SlideId slideId in slideIdList)
{
    var slidePart = presentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
    Console.WriteLine($"Slide {slideIndex + 1}: ID={slideId.Id}, RelationshipId={slideId.RelationshipId}");
    
    // Try to get the title of the slide
    var slide = slidePart.Slide;
    var titleShape = slide.Descendants<P.Shape>()
        .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Title?.Value == "Title 1" ||
                            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Title ||
                            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.CenteredTitle);
    
    if (titleShape != null)
    {
        var titleText = titleShape.TextBody?.Descendants<D.Text>().FirstOrDefault()?.Text;
        if (!string.IsNullOrEmpty(titleText))
        {
            Console.WriteLine($"  Title: {titleText}");
        }
    }
    
    slideIndex++;
}
"""
    
    pptx_file = "/Users/miguelbravo/Downloads/test.pptx"
    
    print(f"Testing with file: {pptx_file}")
    print("Executing C# code...")
    
    result = await execute_csharp_code(test_code, pptx_file)
    
    print("\nResult:")
    if result["success"]:
        print(f"Success! Output:\n{result['output']}")
    else:
        print(f"Failed! Error:\n{result['error']}")
    
    return result

if __name__ == "__main__":
    asyncio.run(test_pptx_editor())