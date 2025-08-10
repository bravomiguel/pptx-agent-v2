using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

public class SemanticAnchor
{
    public string Anchor { get; set; }
    public int Slide { get; set; }
    public string Type { get; set; }
    public string Path { get; set; }
    public string Preview { get; set; }
    public string ParentContext { get; set; }
    public double Confidence { get; set; }
    public Dictionary<string, object> Formatting { get; set; }
    public Dictionary<string, double> Position { get; set; }
}

public class SlideElement
{
    public SemanticAnchor Anchor { get; set; }
    public string Content { get; set; }
    public List<SlideElement> Children { get; set; }
}

public class SlideInfo
{
    public int SlideNumber { get; set; }
    public string Layout { get; set; }
    public string Title { get; set; }
    public List<SlideElement> Elements { get; set; }
}

public class PresentationStructure
{
    public int TotalSlides { get; set; }
    public List<SlideInfo> Slides { get; set; }
}

public static class PptxReader
{
    private static readonly MD5 md5 = MD5.Create();
    
    public static string ReadStructure(string filePath)
    {
        using (var presentation = PresentationDocument.Open(filePath, false))
        {
            var structure = new PresentationStructure
            {
                Slides = new List<SlideInfo>()
            };
            
            var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
            structure.TotalSlides = slideIdList.Count();
            
            int slideNumber = 1;
            foreach (SlideId slideId in slideIdList.OfType<SlideId>())
            {
                var slidePart = (SlidePart)presentation.PresentationPart.GetPartById(slideId.RelationshipId);
                var slideInfo = ProcessSlide(slidePart, slideNumber);
                structure.Slides.Add(slideInfo);
                slideNumber++;
            }
            
            return JsonSerializer.Serialize(structure, new JsonSerializerOptions 
            { 
                WriteIndented = true,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            });
        }
    }
    
    public static string ReadSlideDetails(string filePath, int[] slideNumbers)
    {
        using (var presentation = PresentationDocument.Open(filePath, false))
        {
            var result = new List<SlideInfo>();
            var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
            
            foreach (int slideNum in slideNumbers)
            {
                if (slideNum < 1 || slideNum > slideIdList.Count())
                {
                    continue;
                }
                
                var slideId = slideIdList.OfType<SlideId>().ElementAt(slideNum - 1);
                var slidePart = (SlidePart)presentation.PresentationPart.GetPartById(slideId.RelationshipId);
                var slideInfo = ProcessSlideDetailed(slidePart, slideNum);
                result.Add(slideInfo);
            }
            
            return JsonSerializer.Serialize(result, new JsonSerializerOptions 
            { 
                WriteIndented = true,
                DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
            });
        }
    }
    
    private static SlideInfo ProcessSlide(SlidePart slidePart, int slideNumber)
    {
        var slideInfo = new SlideInfo
        {
            SlideNumber = slideNumber,
            Elements = new List<SlideElement>()
        };
        
        // Get layout name
        if (slidePart.SlideLayoutPart != null)
        {
            var layoutPart = slidePart.SlideLayoutPart;
            var cSld = layoutPart.SlideLayout.CommonSlideData;
            slideInfo.Layout = cSld?.Name?.Value ?? "Unknown";
        }
        
        // Get title
        var titleShape = slidePart.Slide.Descendants<Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Title?.Value == "Title" ||
                                s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.Title ||
                                s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == PlaceholderValues.CenteredTitle);
        
        if (titleShape != null)
        {
            slideInfo.Title = ExtractShapeText(titleShape);
        }
        
        // Process main elements
        int elementIndex = 0;
        foreach (var shape in slidePart.Slide.Descendants<Shape>())
        {
            var element = ProcessShape(shape, slideNumber, elementIndex++, null);
            if (element != null)
            {
                slideInfo.Elements.Add(element);
            }
        }
        
        return slideInfo;
    }
    
    private static SlideInfo ProcessSlideDetailed(SlidePart slidePart, int slideNumber)
    {
        var slideInfo = ProcessSlide(slidePart, slideNumber);
        
        // Add more detailed processing for each element
        foreach (var element in slideInfo.Elements)
        {
            // Add formatting details
            if (element.Anchor.Type == "textbox" || element.Anchor.Type == "bullet")
            {
                // Extract detailed formatting information
                element.Anchor.Formatting = ExtractFormatting(slidePart, element.Anchor.Path);
            }
        }
        
        return slideInfo;
    }
    
    private static SlideElement ProcessShape(Shape shape, int slideNumber, int elementIndex, string parentContext)
    {
        var shapeText = ExtractShapeText(shape);
        if (string.IsNullOrWhiteSpace(shapeText))
            return null;
        
        var element = new SlideElement
        {
            Content = shapeText,
            Children = new List<SlideElement>()
        };
        
        // Determine type
        string elementType = "textbox";
        var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
        if (placeholder != null)
        {
            if (placeholder.Type?.Value == PlaceholderValues.Body)
            {
                elementType = "content";
            }
            else if (placeholder.Type?.Value == PlaceholderValues.Title ||
                     placeholder.Type?.Value == PlaceholderValues.CenteredTitle)
            {
                elementType = "title";
            }
        }
        
        // Check for bullets
        var paragraphs = shape.TextBody?.Descendants<D.Paragraph>();
        if (paragraphs != null && paragraphs.Any(p => p.ParagraphProperties?.GetFirstChild<D.BulletFont>() != null))
        {
            elementType = "bullet";
            
            // Process individual bullets as children
            int bulletIndex = 0;
            foreach (var para in paragraphs)
            {
                var bulletText = ExtractParagraphText(para);
                if (!string.IsNullOrWhiteSpace(bulletText))
                {
                    var bulletElement = new SlideElement
                    {
                        Content = bulletText,
                        Anchor = GenerateAnchor(slideNumber, "bullet", bulletIndex++, bulletText, shapeText)
                    };
                    element.Children.Add(bulletElement);
                }
            }
        }
        
        // Generate anchor
        element.Anchor = GenerateAnchor(slideNumber, elementType, elementIndex, shapeText, parentContext);
        
        // Get position if available
        var transform = shape.ShapeProperties?.Transform2D;
        if (transform != null)
        {
            element.Anchor.Position = new Dictionary<string, double>();
            if (transform.Offset != null)
            {
                element.Anchor.Position["x"] = transform.Offset.X?.Value ?? 0;
                element.Anchor.Position["y"] = transform.Offset.Y?.Value ?? 0;
            }
            if (transform.Extents != null)
            {
                element.Anchor.Position["width"] = transform.Extents.Cx?.Value ?? 0;
                element.Anchor.Position["height"] = transform.Extents.Cy?.Value ?? 0;
            }
        }
        
        return element;
    }
    
    private static SemanticAnchor GenerateAnchor(int slideNumber, string elementType, int elementIndex, string content, string parentContext)
    {
        // Create composite hash
        var hashInput = $"{slideNumber}_{elementType}_{elementIndex}_{parentContext ?? ""}_{content}";
        var hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(hashInput));
        var hash = BitConverter.ToString(hashBytes).Replace("-", "").Substring(0, 6).ToLower();
        
        var anchor = new SemanticAnchor
        {
            Anchor = $"slide{slideNumber}_{elementType}{elementIndex}_{hash}",
            Slide = slideNumber,
            Type = elementType,
            Path = $"slide[{slideNumber}].{elementType}[{elementIndex}]",
            Preview = content.Length > 50 ? content.Substring(0, 50) + "..." : content,
            ParentContext = parentContext,
            Confidence = 1.0
        };
        
        return anchor;
    }
    
    private static string ExtractShapeText(Shape shape)
    {
        if (shape.TextBody == null)
            return null;
        
        var textBuilder = new StringBuilder();
        foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
        {
            var paraText = ExtractParagraphText(paragraph);
            if (!string.IsNullOrWhiteSpace(paraText))
            {
                if (textBuilder.Length > 0)
                    textBuilder.AppendLine();
                textBuilder.Append(paraText);
            }
        }
        
        return textBuilder.ToString();
    }
    
    private static string ExtractParagraphText(D.Paragraph paragraph)
    {
        var textBuilder = new StringBuilder();
        foreach (var run in paragraph.Descendants<D.Run>())
        {
            if (run.Text != null)
            {
                textBuilder.Append(run.Text.Text);
            }
        }
        return textBuilder.ToString();
    }
    
    private static Dictionary<string, object> ExtractFormatting(SlidePart slidePart, string path)
    {
        // Simplified formatting extraction
        var formatting = new Dictionary<string, object>();
        
        // This would be expanded to extract actual formatting details
        // For now, returning basic structure
        formatting["fontSize"] = 14;
        formatting["fontFamily"] = "Arial";
        formatting["bold"] = false;
        formatting["italic"] = false;
        
        return formatting;
    }
    
    public static SlideElement FindByAnchor(PresentationDocument presentation, string anchor)
    {
        // Parse anchor to get slide number
        var parts = anchor.Split('_');
        if (parts.Length < 2 || !parts[0].StartsWith("slide"))
            return null;
        
        if (!int.TryParse(parts[0].Substring(5), out int slideNumber))
            return null;
        
        var slideIdList = presentation.PresentationPart.Presentation.SlideIdList;
        if (slideNumber < 1 || slideNumber > slideIdList.Count())
            return null;
        
        var slideId = slideIdList.OfType<SlideId>().ElementAt(slideNumber - 1);
        var slidePart = (SlidePart)presentation.PresentationPart.GetPartById(slideId.RelationshipId);
        
        // Process slide and find matching anchor
        var slideInfo = ProcessSlideDetailed(slidePart, slideNumber);
        return FindElementByAnchor(slideInfo.Elements, anchor);
    }
    
    private static SlideElement FindElementByAnchor(List<SlideElement> elements, string anchor)
    {
        foreach (var element in elements)
        {
            if (element.Anchor?.Anchor == anchor)
                return element;
            
            if (element.Children != null && element.Children.Any())
            {
                var found = FindElementByAnchor(element.Children, anchor);
                if (found != null)
                    return found;
            }
        }
        return null;
    }
}