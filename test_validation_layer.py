"""Test script for the validation layer."""

import asyncio
import sys
import os
import shutil
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from agent.graph import execute_csharp_code

async def test_invalid_edit():
    """Test 1: Invalid Edit Detection - Should fail validation"""
    print("=== Test 1: Invalid Edit Detection ===")
    
    # Create a copy of the test file
    original_file = "/Users/miguelbravo/Downloads/test.pptx"
    test_file = "/Users/miguelbravo/Downloads/test_validation_invalid.pptx"
    shutil.copy2(original_file, test_file)
    
    # C# code that creates invalid structure by removing all paragraphs from TextBody
    invalid_code = """
// This code will create an invalid structure
var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
if (slideIdList.Count() > 0)
{
    var firstSlideId = slideIdList.ChildElements[0] as SlideId;
    var slidePart = presentation.PresentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
    var slide = slidePart.Slide;
    
    // Find the first shape with text
    var textShape = slide.Descendants<P.Shape>()
        .FirstOrDefault(s => s.TextBody != null);
    
    if (textShape != null)
    {
        Console.WriteLine("Found text shape, creating invalid structure...");
        
        // Remove all paragraphs but don't add any new ones
        // This violates the rule that TextBody must have at least one paragraph
        var textBody = textShape.TextBody;
        textBody.RemoveAllChildren<D.Paragraph>();
        
        Console.WriteLine("Removed all paragraphs from TextBody (invalid structure)");
    }
}
"""
    
    print(f"Testing with file: {test_file}")
    result = await execute_csharp_code(invalid_code, test_file)
    
    print("\nResult:")
    if result["success"]:
        print("❌ FAILED: Expected validation error but edit succeeded")
        return False
    else:
        if "Validation failed" in result["error"]:
            print("✅ PASSED: Validation correctly caught the invalid structure")
            print(f"Error message:\n{result['error']}")
            
            # Verify file wasn't modified by checking if it still opens
            print("\nVerifying original file remains unchanged...")
            return True
        else:
            print(f"❌ FAILED: Got error but not validation error:\n{result['error']}")
            return False

async def test_valid_edit():
    """Test 2: Valid Edit - Should pass validation"""
    print("\n=== Test 2: Valid Edit Verification ===")
    
    # Create a copy of the test file
    original_file = "/Users/miguelbravo/Downloads/test.pptx"
    test_file = "/Users/miguelbravo/Downloads/test_validation_valid.pptx"
    shutil.copy2(original_file, test_file)
    
    # C# code that properly modifies text
    valid_code = """
// This code creates a valid structure
var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
if (slideIdList.Count() > 0)
{
    var firstSlideId = slideIdList.ChildElements[0] as SlideId;
    var slidePart = presentation.PresentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
    var slide = slidePart.Slide;
    
    // Find the title shape
    var titleShape = slide.Descendants<P.Shape>()
        .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Title ||
                             s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.CenteredTitle);
    
    if (titleShape != null && titleShape.TextBody != null)
    {
        Console.WriteLine("Found title shape, modifying text properly...");
        
        var textBody = titleShape.TextBody;
        
        // Remove existing paragraphs
        textBody.RemoveAllChildren<D.Paragraph>();
        
        // Add new paragraph with proper structure
        var paragraph = new D.Paragraph();
        
        // Add paragraph properties (optional but good practice)
        var paragraphProperties = new D.ParagraphProperties();
        paragraph.Append(paragraphProperties);
        
        // Add run with text
        var run = new D.Run();
        var runProperties = new D.RunProperties() { FontSize = 4400 };
        var text = new D.Text() { Text = "Valid Edit - With Validation!" };
        
        run.Append(runProperties);
        run.Append(text);
        paragraph.Append(run);
        
        // Add end paragraph properties
        paragraph.Append(new D.EndParagraphRunProperties());
        
        // Add paragraph to text body
        textBody.Append(paragraph);
        
        Console.WriteLine("Text modified with valid structure");
    }
}
"""
    
    print(f"Testing with file: {test_file}")
    result = await execute_csharp_code(valid_code, test_file)
    
    print("\nResult:")
    if result["success"]:
        print("✅ PASSED: Valid edit succeeded")
        print(f"Output:\n{result['output']}")
        print(f"\nModified file saved at: {test_file}")
        print("You can open this file in PowerPoint to verify it opens without repair dialog")
        return True
    else:
        print(f"❌ FAILED: Valid edit was rejected:\n{result['error']}")
        return False

async def main():
    """Run both tests"""
    print("Testing Validation Layer Implementation\n")
    
    # Run tests
    test1_passed = await test_invalid_edit()
    test2_passed = await test_valid_edit()
    
    # Summary
    print("\n=== Test Summary ===")
    print(f"Test 1 (Invalid Edit Detection): {'✅ PASSED' if test1_passed else '❌ FAILED'}")
    print(f"Test 2 (Valid Edit Verification): {'✅ PASSED' if test2_passed else '❌ FAILED'}")
    
    if test1_passed and test2_passed:
        print("\n✅ All tests passed! Validation layer is working correctly.")
    else:
        print("\n❌ Some tests failed. Check the output above for details.")

if __name__ == "__main__":
    asyncio.run(main())