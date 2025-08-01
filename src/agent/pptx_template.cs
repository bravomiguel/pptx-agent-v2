using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

public class PptxEditor
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
                {CODE}
                // USER_CODE_END
                
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