using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

public class TestReader
{
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: TestReader <pptx-file> <command>");
            Console.WriteLine("Commands: structure, details <slide-numbers>, test-anchors");
            Environment.Exit(1);
        }
        
        string filePath = args[0];
        string command = args[1];
        
        try
        {
            switch (command.ToLower())
            {
                case "structure":
                    Console.WriteLine("=== PRESENTATION STRUCTURE ===");
                    string structure = PptxReader.ReadStructure(filePath);
                    Console.WriteLine(structure);
                    break;
                    
                case "details":
                    if (args.Length < 3)
                    {
                        Console.WriteLine("Error: Please provide slide numbers (e.g., 1,2,3)");
                        Environment.Exit(1);
                    }
                    int[] slideNumbers = args[2].Split(',').Select(int.Parse).ToArray();
                    Console.WriteLine($"=== SLIDE DETAILS FOR SLIDES {string.Join(", ", slideNumbers)} ===");
                    string details = PptxReader.ReadSlideDetails(filePath, slideNumbers);
                    Console.WriteLine(details);
                    break;
                    
                case "test-anchors":
                    Console.WriteLine("=== TESTING ANCHOR SYSTEM ===");
                    TestAnchorSystem(filePath);
                    break;
                    
                default:
                    Console.WriteLine($"Unknown command: {command}");
                    Environment.Exit(1);
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
            Environment.Exit(1);
        }
    }
    
    private static void TestAnchorSystem(string filePath)
    {
        // Read structure to get anchors
        Console.WriteLine("1. Reading presentation structure...");
        string structure = PptxReader.ReadStructure(filePath);
        
        // Parse JSON to extract first anchor (simplified for testing)
        var lines = structure.Split('\n');
        string firstAnchor = null;
        foreach (var line in lines)
        {
            if (line.Contains("\"Anchor\":"))
            {
                var start = line.IndexOf(": \"") + 3;
                var end = line.LastIndexOf("\"");
                if (start > 2 && end > start)
                {
                    firstAnchor = line.Substring(start, end - start);
                    break;
                }
            }
        }
        
        if (firstAnchor != null)
        {
            Console.WriteLine($"2. Found anchor: {firstAnchor}");
            
            // Test finding element by anchor
            using (var presentation = PresentationDocument.Open(filePath, false))
            {
                var element = PptxReader.FindByAnchor(presentation, firstAnchor);
                if (element != null)
                {
                    Console.WriteLine($"3. Successfully found element with anchor!");
                    Console.WriteLine($"   Content: {element.Content}");
                    Console.WriteLine($"   Type: {element.Anchor.Type}");
                    Console.WriteLine($"   Slide: {element.Anchor.Slide}");
                }
                else
                {
                    Console.WriteLine("3. Could not find element by anchor");
                }
            }
        }
        else
        {
            Console.WriteLine("No anchors found in presentation");
        }
        
        Console.WriteLine("\n4. Testing collision prevention:");
        // Read all slides to check for unique anchors
        using (var presentation = PresentationDocument.Open(filePath, false))
        {
            var slideCount = presentation.PresentationPart.Presentation.SlideIdList.Count();
            int[] allSlides = Enumerable.Range(1, slideCount).ToArray();
            string allDetails = PptxReader.ReadSlideDetails(filePath, allSlides);
            
            var anchors = new System.Collections.Generic.HashSet<string>();
            int duplicates = 0;
            
            foreach (var line in allDetails.Split('\n'))
            {
                if (line.Contains("\"Anchor\":"))
                {
                    var start = line.IndexOf(": \"") + 3;
                    var end = line.LastIndexOf("\"");
                    if (start > 2 && end > start)
                    {
                        var anchor = line.Substring(start, end - start);
                        if (!anchors.Add(anchor))
                        {
                            duplicates++;
                            Console.WriteLine($"   Duplicate anchor found: {anchor}");
                        }
                    }
                }
            }
            
            Console.WriteLine($"   Total unique anchors: {anchors.Count}");
            Console.WriteLine($"   Duplicate anchors: {duplicates}");
            
            if (duplicates == 0)
            {
                Console.WriteLine("   ✓ All anchors are unique!");
            }
        }
    }
}