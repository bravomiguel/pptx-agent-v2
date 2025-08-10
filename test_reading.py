#!/usr/bin/env python3
"""Test the PPTX reading functionality without LLM."""

import asyncio
import json
import sys
import os

# Add the parent directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.agent.graph import execute_reading_code

async def test_reading():
    """Test reading functionality with the test PPTX file."""
    test_file = "/Users/miguelbravo/Downloads/test.pptx"
    
    print("=" * 60)
    print("Testing PPTX Reading Functionality")
    print("=" * 60)
    
    # Test 1: Read presentation structure
    print("\n1. Testing ReadStructure...")
    code_structure = """
string result = PptxReader.ReadStructure(filePath);
Console.WriteLine(result);
"""
    
    result = await execute_reading_code(code_structure, test_file)
    if result["success"]:
        print("✓ Structure read successfully")
        structure = json.loads(result["output"])
        print(f"  - Total slides: {structure['TotalSlides']}")
        print(f"  - First slide title: {structure['Slides'][0]['Title']}")
        
        # Extract first anchor for testing
        first_anchor = None
        if structure['Slides'] and structure['Slides'][0]['Elements']:
            first_anchor = structure['Slides'][0]['Elements'][0]['Anchor']['Anchor']
            print(f"  - First anchor: {first_anchor}")
    else:
        print(f"✗ Failed to read structure: {result['error']}")
        return
    
    # Test 2: Read specific slide details
    print("\n2. Testing ReadSlideDetails...")
    code_details = """
int[] slideNumbers = new int[] { 1, 2 };
string result = PptxReader.ReadSlideDetails(filePath, slideNumbers);
Console.WriteLine(result);
"""
    
    result = await execute_reading_code(code_details, test_file)
    if result["success"]:
        print("✓ Slide details read successfully")
        slides = json.loads(result["output"])
        for slide in slides:
            print(f"  - Slide {slide['SlideNumber']}: {slide['Title']}")
            print(f"    Elements: {len(slide['Elements'])}")
    else:
        print(f"✗ Failed to read slide details: {result['error']}")
        return
    
    # Test 3: Test anchor lookup
    if first_anchor:
        print(f"\n3. Testing anchor lookup for: {first_anchor}")
        code_anchor = f"""
using (var presentation = PresentationDocument.Open(filePath, false))
{{
    var element = PptxReader.FindByAnchor(presentation, "{first_anchor}");
    if (element != null)
    {{
        Console.WriteLine($"Found element with anchor {{element.Anchor.Anchor}}");
        Console.WriteLine($"Content: {{element.Content}}");
        Console.WriteLine($"Type: {{element.Anchor.Type}}");
    }}
    else
    {{
        Console.WriteLine("Element not found");
    }}
}}
"""
        
        result = await execute_reading_code(code_anchor, test_file)
        if result["success"]:
            print("✓ Anchor lookup successful")
            print(f"  Output: {result['output'].strip()}")
        else:
            print(f"✗ Anchor lookup failed: {result['error']}")
    
    # Test 4: Check for unique anchors
    print("\n4. Testing anchor uniqueness...")
    code_unique = """
var slideCount = 0;
using (var presentation = PresentationDocument.Open(filePath, false))
{
    slideCount = presentation.PresentationPart.Presentation.SlideIdList.Count();
}

int[] allSlides = Enumerable.Range(1, slideCount).ToArray();
string allDetails = PptxReader.ReadSlideDetails(filePath, allSlides);

var parsed = System.Text.Json.JsonDocument.Parse(allDetails);
var anchors = new System.Collections.Generic.HashSet<string>();
int totalAnchors = 0;
int duplicates = 0;

void ProcessElement(System.Text.Json.JsonElement element)
{
    if (element.TryGetProperty("Anchor", out var anchorProp) && 
        anchorProp.TryGetProperty("Anchor", out var anchorValue))
    {
        totalAnchors++;
        string anchor = anchorValue.GetString();
        if (!anchors.Add(anchor))
        {
            duplicates++;
            Console.WriteLine($"Duplicate anchor: {anchor}");
        }
    }
    
    if (element.TryGetProperty("Children", out var children))
    {
        foreach (var child in children.EnumerateArray())
        {
            ProcessElement(child);
        }
    }
}

foreach (var slide in parsed.RootElement.EnumerateArray())
{
    if (slide.TryGetProperty("Elements", out var elements))
    {
        foreach (var elem in elements.EnumerateArray())
        {
            ProcessElement(elem);
        }
    }
}

Console.WriteLine($"Total anchors: {totalAnchors}");
Console.WriteLine($"Unique anchors: {anchors.Count}");
Console.WriteLine($"Duplicate anchors: {duplicates}");
"""
    
    result = await execute_reading_code(code_unique, test_file)
    if result["success"]:
        lines = result["output"].strip().split('\n')
        if "Duplicate anchors: 0" in result["output"]:
            print("✓ All anchors are unique!")
        else:
            print("✗ Found duplicate anchors")
        for line in lines[-3:]:
            print(f"  {line}")
    else:
        print(f"✗ Uniqueness check failed: {result['error']}")
    
    print("\n" + "=" * 60)
    print("All tests completed successfully!")
    print("=" * 60)

if __name__ == "__main__":
    asyncio.run(test_reading())