using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Error: Please provide the PowerPoint file path");
            Environment.Exit(1);
        }
        
        string filePath = args[0];
        
        try
        {
            using (PresentationDocument presentation = PresentationDocument.Open(filePath, true))
            {
                // USER_CODE_START
                // {CODE}
                // USER_CODE_END
                
                // Validate the presentation before saving
                // We'll validate only slide parts to avoid pre-existing chart issues
                var validator = new OpenXmlValidator();
                var hasErrors = false;
                
                // Validate each slide part
                foreach (var slidePart in presentation.PresentationPart.SlideParts)
                {
                    var slideErrors = validator.Validate(slidePart).ToList();
                    if (slideErrors.Any())
                    {
                        if (!hasErrors)
                        {
                            Console.WriteLine("VALIDATION_ERROR: The presentation has structural errors:");
                            hasErrors = true;
                        }
                        
                        foreach (var error in slideErrors)
                        {
                            var path = error.Path?.XPath ?? "Unknown path";
                            var part = error.Part?.Uri?.ToString() ?? "Unknown part";
                            Console.WriteLine($"- [{error.ErrorType}] in {part} at {path}: {error.Description}");
                            if (!string.IsNullOrEmpty(error.Id))
                            {
                                Console.WriteLine($"  Rule ID: {error.Id}");
                            }
                        }
                    }
                }
                
                if (hasErrors)
                {
                    Environment.Exit(2); // Special exit code for validation failure
                }
                
                Console.WriteLine("Successfully executed PowerPoint modifications");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Environment.Exit(1);
        }
    }
}