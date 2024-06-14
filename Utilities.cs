using System;
using System.Xml.Linq;
using WorksheetGenerator.Elements;

namespace WorksheetGenerator.Utilities
{
    public static class HF
    {
        public static XElement? GetElementOnly(XElement? element)
        {
            if (element != null)
            {
                // Create a new element with the same name and attributes
                XElement newElement = new XElement(element.Name);
                foreach (XAttribute attribute in element.Attributes())
                    newElement.Add(new XAttribute(attribute));
                return newElement;
            }
            else
                return null;
        }

        public static XElement? GetDocumentAndBodyOnly(XDocument doc)
        {
            XElement? originalDocElement = doc.Element(El.w + "document");
            XElement? docElement = GetElementOnly(doc.Element(El.w + "document"));
            if (docElement != null && originalDocElement != null)
            {
                XElement? body = GetElementOnly(originalDocElement.Element(El.w + "body"));
                if (body != null)
                {
                    docElement.Add(body);
                    return docElement;
                }
                else
                    return null;
            }
            else
                return null;
        }

        public static bool IsIdentifier(XElement element)
        {
            XElement? runElement = element.Descendants(El.w + "r").FirstOrDefault();
            XElement? textElement = element.Descendants(El.w + "t").FirstOrDefault();
            if (textElement != null && runElement != null)
            {
                // Check if the text is all uppercase
                bool isUpperCase = textElement.Value.All(char.IsUpper);

                // Check if the text is bold
                XElement? runProperties = runElement.Element(El.w + "rPr");
                if (runProperties != null)
                {
                    bool isBold = runProperties.Element(El.w + "b") != null;
                    return isBold && isUpperCase;
                }
                else
                    return false;
            }
            else
                return false;
        }

        public static bool IsWhiteSpaceOnly(XElement element)
        {
            // Get the text within the <w:p> element
            string? text = element.Descendants()
                                .Where(e => e.Name.LocalName == "t")
                                .Select(e => e.Value)
                                .FirstOrDefault();

            // Check if the text is null or consists only of whitespace characters
            return string.IsNullOrWhiteSpace(text);
        }

        public static bool StartsWith(XElement element, string str)
        {
            return ((string)element).Trim().ToLower().StartsWith(str.ToLower());
        }

        public static string RemovePrefix(string text)
        {
            // Find the index of the colon
            int colonIndex = text.IndexOf(':');

            // If colon is found and it's not the last character
            if (colonIndex != -1 && colonIndex < text.Length - 1)
            {
                // Extract and return the substring after the colon (excluding colon and leading whitespace)
                return text.Substring(colonIndex + 1).TrimStart();
            }

            // Return the original string if colon is not found or it's the last character
            return text;
        }

        public static XElement? GetWorksheetTitleElement(IEnumerable<XElement> paragraphs)
        {
            foreach (XElement paragraph in paragraphs)
            {
                if (StartsWith(paragraph, "title:"))
                {
                    string title = RemovePrefix((string)paragraph).ToUpper();
                    XElement worksheetTitleElement = El.Paragraph(title);
                    AddWorksheetTitleStyles(worksheetTitleElement);
                    return worksheetTitleElement;
                }
            }
            return null;
        }

        public static List<XElement> GetParagraphsByIdentifier(IEnumerable<XElement> paragraphs, string identifierName)
        {
            bool isBetweenIdentifiers = false;
            List<XElement> result = new List<XElement>();

            foreach (XElement paragraph in paragraphs)
            {
                if (IsIdentifier(paragraph))
                {
                    if (isBetweenIdentifiers)
                        break;

                    if (((string)paragraph).Trim().StartsWith(identifierName))
                    {
                        isBetweenIdentifiers = true;
                        result.Clear();
                    }
                }
                else if (isBetweenIdentifiers)
                    if (!StartsWith(paragraph, "chatgpt:"))
                        result.Add(paragraph);
            }

            return result;
        }

        public static bool IsImage(XElement element)
        {
            foreach (XElement child in element.Elements())
            {
                // Check if the child element is a drawing or picture
                if (child.Name == El.w + "drawing" || child.Name == El.w + "pict")
                    return true;

                // Check if the child element is a run element containing drawing or picture
                if (child.Name == El.w + "r")
                    foreach (XElement runChild in child.Elements())
                        if (runChild.Name == El.w + "drawing" || runChild.Name == El.w + "pict")
                            return true;
            }

            return false;
        }

        public static ulong GetWidth(double width, double height, double desiredHeight)
        {
            return (ulong)Math.Round((double)(desiredHeight / height * width));
        }

        public static void FormatImage(XElement element, double desiredHeight = 1640000)
        {
            // Add elements to <w:pPr>
            XElement? paragraphProperty = element.Element(El.w + "pPr");
            paragraphProperty?.AddFirst(new XElement(El.w + "jc",
                new XAttribute(El.w + "val", "center")
            ));
            paragraphProperty?.AddFirst(new XElement(El.w + "spacing",
                new XAttribute(El.w + "before", 240),
                new XAttribute(El.w + "line", 400),
                new XAttribute(El.w + "lineRule", "auto")
            ));

            // Add elements to <w:rPr>
            XElement? runProperty = paragraphProperty?.Element(El.w + "rPr");
            runProperty?.AddFirst(new XElement(El.w + "szCs",
                new XAttribute(El.w + "val", 36)
            ));
            runProperty?.AddFirst(new XElement(El.w + "sz",
                new XAttribute(El.w + "val", 36)
            ));
            runProperty?.AddFirst(new XElement(El.w + "bCs"));
            runProperty?.AddFirst(new XElement(El.w + "b"));

            // Edit <wp:extent> (Resizing)
            XElement? extentElement = element.Descendants(El.wp + "extent").FirstOrDefault();
            XAttribute? cx = extentElement?.Attribute("cx");
            XAttribute? cy = extentElement?.Attribute("cy");
            if (cx != null && cy != null)
            {
                bool validCx = double.TryParse(cx.Value, out double origWidth);
                bool validCy = double.TryParse(cy.Value, out double origHeight);

                if (validCx && validCy)
                {
                    double desiredWidth = GetWidth(origWidth, origHeight, desiredHeight);
                    extentElement?.SetAttributeValue("cx", desiredWidth);
                    extentElement?.SetAttributeValue("cy", desiredHeight);

                    // Edit <a:xfrm><a:ext> (Resizing)
                    XElement? transformElement = element.Descendants(El.a + "xfrm").FirstOrDefault();
                    XElement? extElement = transformElement?.Element(El.a + "ext");

                    extElement?.SetAttributeValue("cx", desiredWidth);
                    extElement?.SetAttributeValue("cy", desiredHeight);
                }
            }

            // Edit <a:prstGeom> (Rounded corners)
            XElement? geometryElement = element.Descendants(El.a + "prstGeom").FirstOrDefault();
            geometryElement?.SetAttributeValue("prst", "roundRect");
        }

        public static void AddSectionTitleStyles(XElement paragraph)
        {
            El.CenterParagraph(paragraph);
            El.SetParagraphSize(paragraph, 36);
            El.AddBoldToParagraph(paragraph);
            El.SetParagraphColor(paragraph, "0F9ED5", "accent4");
        }

        public static void AddWorksheetTitleStyles(XElement paragraph)
        {
            El.CenterParagraph(paragraph);
            El.SetParagraphSize(paragraph, 48);
            El.AddBoldToParagraph(paragraph);
        }

        public static XElement? GetSectionTitleElement(IEnumerable<XElement> elements)
        {
            foreach (XElement element in elements)
                if (StartsWith(element, "title:"))
                    return element;

            return null;
        }

        public static XElement GetFormattedSectionTitleElement(string title, int sectionNo = -1)
        {
            string formattedTitle;
            if (sectionNo >= 0)
                formattedTitle = sectionNo + ". " + title.Trim();
            else
                formattedTitle = title.Trim();

            XElement titleElement = El.Paragraph(formattedTitle);
            AddSectionTitleStyles(titleElement);
            return titleElement;
        }

        public static Dictionary<string, string> GetVocab(List<XElement> paragraphs)
        {
            Dictionary<string, string> vocab = [];

            foreach (var paragraph in paragraphs)
                Console.WriteLine((string)paragraph);

            return vocab;
        }

        public static List<XElement> GetProcessedVocab(IEnumerable<XElement> allParagraphs, int sectionNo = -1)
        {
            List<XElement> paragraphs = GetParagraphsByIdentifier(allParagraphs, "VOCAB");
            List<XElement> result = [];

            // Format & add title
            XElement titleElement = GetFormattedSectionTitleElement("Vocabulary", sectionNo);
            result.Add(titleElement);

            // Vocab box
            Dictionary<string, string> vocab = GetVocab(paragraphs);

            return result;
        }

        public static List<XElement> GetProcessedReading(IEnumerable<XElement> allParagraphs, int sectionNo = -1)
        {
            List<XElement> paragraphs = GetParagraphsByIdentifier(allParagraphs, "READING");
            List<XElement> result = [];
            XElement? origTitleElement = GetSectionTitleElement(paragraphs);

            // Format & add title
            if (origTitleElement != null)
            {
                XElement newTitleElement = GetFormattedSectionTitleElement(
                    RemovePrefix((string)origTitleElement),
                    sectionNo
                );
                result.Add(newTitleElement);
            }

            // Get main passage paragraphs
            List<XElement> passageParagraphs = [];
            List<XElement> previewImages = [];
            bool isBeforePassage = true;
            foreach (XElement paragraph in paragraphs)
            {
                if (!StartsWith(paragraph, "title:"))
                {
                    // Format any images
                    if (IsImage(paragraph))
                        FormatImage(paragraph);
                    else if (isBeforePassage)
                        isBeforePassage = false;
                    if (isBeforePassage)
                        previewImages.Add(paragraph);
                    else if (!IsWhiteSpaceOnly(paragraph) || IsImage(paragraph))
                        passageParagraphs.Add(paragraph);
                }
            }

            // Add preview images
            foreach (XElement element in previewImages)
                result.Add(element);

            // Create table for main passage & add passage paragraphs
            XElement tableElement = El.TableElement(2, [6374, 2976], [passageParagraphs, null]);
            result.Add(tableElement);

            return result;
        }
    }
}