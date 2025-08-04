"""Test that the C# code execution works with the corrected format."""

import asyncio
from src.agent.graph import execute_pptx_code, State


async def test_csharp_execution():
    """Test the tool with properly formatted C# code (no using statements, no return)."""
    
    # Create a mock state with pptx_file_path
    mock_state = State(
        messages=[],
        pptx_file_path="/Users/miguelbravo/Downloads/test.pptx"
    )
    
    # Corrected C# code - no using statements, no return, just Console.WriteLine
    corrected_code = """
// Add a new slide with the title "Hello world!" using the default title layout
try
{
    var presentationPart = presentation.PresentationPart;
    if (presentationPart == null)
        throw new Exception("PresentationPart is missing.");

    // Get the first slide master and a title layout
    var slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
    if (slideMasterPart == null)
        throw new Exception("No SlideMasterPart found.");

    var titleLayoutPart = slideMasterPart.SlideLayoutParts.FirstOrDefault(
        sl => sl.SlideLayout.CommonSlideData.Name?.Value?.ToLower().Contains("title") ?? false
    );
    if (titleLayoutPart == null)
        titleLayoutPart = slideMasterPart.SlideLayoutParts.FirstOrDefault(); // fallback
    if (titleLayoutPart == null)
        throw new Exception("No SlideLayoutPart found.");

    // Create a new slide part
    var newSlidePart = presentationPart.AddNewPart<SlidePart>();
    
    // Create a new slide with minimal required structure
    var slide = new Slide(
        new CommonSlideData(
            new ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = 1, Name = "" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()
                ),
                new P.GroupShapeProperties(
                    new D.TransformGroup(),
                    new D.Extents() { Cx = 0, Cy = 0 },
                    new D.Offset() { X = 0, Y = 0 },
                    new D.ChildExtents() { Cx = 0, Cy = 0 },
                    new D.ChildOffset() { X = 0, Y = 0 }
                ),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 2, Name = "Title" },
                        new P.NonVisualShapeDrawingProperties(
                            new D.ShapeLocks() { NoGrouping = true }
                        ),
                        new P.ApplicationNonVisualDrawingProperties(
                            new P.PlaceholderShape() { Type = PlaceholderValues.Title }
                        )
                    ),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new D.BodyProperties(),
                        new D.ListStyle(),
                        new D.Paragraph(
                            new D.Run(
                                new D.Text("Hello world!")
                            )
                        )
                    )
                )
            )
        ),
        new ColorMapOverride(
            new D.MasterColorMapping()
        )
    );
    
    newSlidePart.Slide = slide;
    newSlidePart.Slide.Save();

    // Add the new slide to the slide list
    var slideIdList = presentationPart.Presentation.SlideIdList;
    uint maxSlideId = 256;
    if (slideIdList.ChildElements.Count > 0)
    {
        maxSlideId = slideIdList.ChildElements
            .OfType<SlideId>()
            .Max(s => s.Id.Value) + 1;
    }
    var relId = presentationPart.GetIdOfPart(newSlidePart);
    slideIdList.Append(new SlideId() { Id = maxSlideId, RelationshipId = relId });
    presentationPart.Presentation.Save();
    
    Console.WriteLine("A new slide titled 'Hello world!' has been added to the end of your presentation.");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}
"""
    
    print("Testing C# code execution with corrected format...")
    print("PowerPoint file:", mock_state.pptx_file_path)
    
    try:
        # Call the tool directly
        result = await execute_pptx_code.ainvoke({
            "code": corrected_code,
            "state": mock_state
        })
        
        print("\nResult:", result)
        
        if "successfully" in result.lower():
            print("✅ C# code executed successfully!")
        else:
            print("❌ Execution failed, but this might be expected for complex operations")
            
        # Test simple code that should definitely work
        simple_code = """
var slideCount = presentation.PresentationPart.Presentation.SlideIdList.ChildElements.Count;
Console.WriteLine($"Presentation has {slideCount} slides");
"""
        
        print("\n\nTesting simple C# code...")
        result2 = await execute_pptx_code.ainvoke({
            "code": simple_code,
            "state": mock_state
        })
        
        print("Result:", result2)
        
        if "successfully" in result2.lower():
            print("✅ Simple test passed!")
        else:
            print("❌ Simple test failed")
            
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        raise


if __name__ == "__main__":
    asyncio.run(test_csharp_execution())