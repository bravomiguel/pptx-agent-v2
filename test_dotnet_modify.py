"""Test script that modifies a PowerPoint presentation."""

import asyncio
import sys
import os
import shutil
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from agent.graph import execute_csharp_code

async def test_pptx_modification():
    """Test the PowerPoint editor by modifying a presentation."""
    
    # Create a copy of the test file to modify
    original_file = "/Users/miguelbravo/Downloads/test.pptx"
    test_file = "/Users/miguelbravo/Downloads/test_modified.pptx"
    shutil.copy2(original_file, test_file)
    print(f"Created copy of test file: {test_file}")
    
    # C# code that modifies the first slide's title
    modify_code = """
// Modify the title of the first slide
var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
if (slideIdList.Count() > 0)
{
    var firstSlideId = slideIdList.ChildElements[0] as SlideId;
    var slidePart = presentation.PresentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
    var slide = slidePart.Slide;
    
    // Find the title shape
    var titleShape = slide.Descendants<P.Shape>()
        .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Title?.Value == "Title 1" ||
                            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Title ||
                            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.CenteredTitle);
    
    if (titleShape != null)
    {
        Console.WriteLine("Found title shape, modifying text...");
        
        // Clear existing text
        var textBody = titleShape.TextBody;
        if (textBody != null)
        {
            textBody.RemoveAllChildren<D.Paragraph>();
            
            // Add new paragraph with text
            var paragraph = new D.Paragraph();
            var run = new D.Run();
            var text = new D.Text() { Text = "Modified by .NET Code!" };
            
            // Add run properties for formatting
            var runProperties = new D.RunProperties()
            {
                FontSize = 4400, // 44pt
                Bold = true
            };
            
            run.Append(runProperties);
            run.Append(text);
            paragraph.Append(run);
            textBody.Append(paragraph);
            
            Console.WriteLine("Title text modified successfully!");
        }
    }
    else
    {
        Console.WriteLine("Title shape not found on first slide");
    }
}
else
{
    Console.WriteLine("No slides found in presentation");
}
"""
    
    print("\nExecuting modification code...")
    result = await execute_csharp_code(modify_code, test_file)
    
    print("\nResult:")
    if result["success"]:
        print(f"Success! Output:\n{result['output']}")
        print(f"\nModified file saved at: {test_file}")
    else:
        print(f"Failed! Error:\n{result['error']}")
    
    return result

if __name__ == "__main__":
    asyncio.run(test_pptx_modification())