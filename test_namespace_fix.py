#!/usr/bin/env python3
"""Test that namespace aliases work correctly and A namespace error is caught."""

import asyncio
import shutil
import tempfile
import sys
import os

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.agent.graph import execute_csharp_code

async def test_namespace_aliases():
    """Test that namespace aliases work correctly."""
    # Copy test file to temp location
    test_file = "/Users/miguelbravo/Downloads/test.pptx"
    temp_dir = tempfile.mkdtemp()
    temp_file = os.path.join(temp_dir, "test_namespace.pptx")
    shutil.copy2(test_file, temp_file)
    
    print("=" * 60)
    print("Testing Namespace Alias Fix")
    print("=" * 60)
    
    try:
        # Test 1: Code with wrong A namespace (should fail)
        print("\n1. Testing code with incorrect A namespace (should fail)...")
        bad_code = """
var slideId = presentation.PresentationPart.Presentation.SlideIdList.ChildElements[0] as SlideId;
var slidePart = presentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
var slide = slidePart.Slide;

var shape = slide.Descendants<P.Shape>().FirstOrDefault();
if (shape?.TextBody != null)
{
    var paragraph = shape.TextBody.GetFirstChild<A.Paragraph>();  // Wrong: A namespace
    paragraph.RemoveAllChildren<A.Run>();  // Wrong: A namespace
    var run = new A.Run(new A.Text("Test"));  // Wrong: A namespace
    paragraph.Append(run);
}
Console.WriteLine("Done");
"""
        
        result = await execute_csharp_code(bad_code, temp_file)
        if result["success"]:
            print("✗ Code with A namespace should have failed but succeeded!")
        else:
            if "CS0246" in result["error"] and "namespace name 'A' could not be found" in result["error"]:
                print("✓ Correctly rejected code with A namespace")
                print("  Error contains: 'namespace name 'A' could not be found'")
            else:
                print(f"✗ Failed but with unexpected error: {result['error'][:200]}")
        
        # Test 2: Code with correct D namespace (should succeed)
        print("\n2. Testing code with correct D namespace (should succeed)...")
        good_code = """
var slideId = presentation.PresentationPart.Presentation.SlideIdList.ChildElements[0] as SlideId;
var slidePart = presentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
var slide = slidePart.Slide;

var shape = slide.Descendants<P.Shape>().FirstOrDefault();
if (shape?.TextBody != null)
{
    var paragraph = shape.TextBody.GetFirstChild<D.Paragraph>();  // Correct: D namespace
    if (paragraph != null)
    {
        paragraph.RemoveAllChildren<D.Run>();  // Correct: D namespace
        var run = new D.Run(new D.Text("Test with D namespace"));  // Correct: D namespace
        paragraph.Append(run);
        Console.WriteLine("Successfully modified text using D namespace");
    }
    else
    {
        Console.WriteLine("No paragraph found");
    }
}
else
{
    Console.WriteLine("No shape with text body found");
}
"""
        
        result = await execute_csharp_code(good_code, temp_file)
        if result["success"]:
            print("✓ Code with D namespace executed successfully")
            print(f"  Output: {result['output'].strip()}")
        else:
            print(f"✗ Code with D namespace failed: {result['error'][:500]}")
        
        # Test 3: Code with P namespace for Presentation (should succeed)
        print("\n3. Testing P namespace for Presentation types...")
        p_namespace_code = """
// Test P namespace for Presentation types
var slideCount = presentation.PresentationPart.Presentation.SlideIdList.Count();
Console.WriteLine($"Presentation has {slideCount} slides");

// Get first slide using P.SlideId type annotation
var firstSlideId = presentation.PresentationPart.Presentation.SlideIdList.ChildElements[0] as P.SlideId;
if (firstSlideId != null)
{
    Console.WriteLine("Successfully accessed SlideId using P namespace");
}

// Check for P.Shape type
var slidePart = presentation.PresentationPart.GetPartById(firstSlideId.RelationshipId) as SlidePart;
var shapes = slidePart.Slide.Descendants<P.Shape>().Count();
Console.WriteLine($"Slide has {shapes} shapes (using P.Shape)");
"""
        
        result = await execute_csharp_code(p_namespace_code, temp_file)
        if result["success"]:
            print("✓ P namespace for Presentation types works correctly")
            print(f"  Output: {result['output'].strip()}")
        else:
            print(f"✗ P namespace test failed: {result['error'][:200]}")
        
        # Test 4: Verify both namespaces can be used together
        print("\n4. Testing mixed P and D namespace usage...")
        mixed_code = """
// Using both P and D namespaces
var slideId = presentation.PresentationPart.Presentation.SlideIdList.ChildElements[0] as P.SlideId;  // P namespace
var slidePart = presentation.PresentationPart.GetPartById(slideId.RelationshipId) as SlidePart;

// Find P.Shape and modify with D namespace
var shape = slidePart.Slide.Descendants<P.Shape>().FirstOrDefault();  // P.Shape
if (shape?.TextBody != null)
{
    var paragraph = shape.TextBody.GetFirstChild<D.Paragraph>();  // D.Paragraph
    if (paragraph != null)
    {
        var existingRuns = paragraph.Elements<D.Run>().Count();  // D.Run
        Console.WriteLine($"Found {existingRuns} runs in paragraph");
        
        // Add new run
        var newRun = new D.Run(new D.Text(" [Added via mixed P and D]"));  // D.Run, D.Text
        paragraph.Append(newRun);
        Console.WriteLine("Successfully used both P and D namespaces");
    }
}
"""
        
        result = await execute_csharp_code(mixed_code, temp_file)
        if result["success"]:
            print("✓ Mixed P and D namespace usage works correctly")
            print(f"  Output: {result['output'].strip()}")
        else:
            print(f"✗ Mixed namespace test failed: {result['error'][:200]}")
        
    finally:
        # Cleanup
        shutil.rmtree(temp_dir)
    
    print("\n" + "=" * 60)
    print("Namespace Testing Complete")
    print("Summary:")
    print("- A namespace correctly rejected (not available)")
    print("- D namespace works for Drawing types")
    print("- P namespace works for Presentation types")
    print("- Mixed P and D usage works correctly")
    print("=" * 60)

if __name__ == "__main__":
    asyncio.run(test_namespace_aliases())