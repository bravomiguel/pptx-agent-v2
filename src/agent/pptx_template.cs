#r "nuget: DocumentFormat.OpenXml, 3.0.0"

using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

string filePath = Args[0];

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