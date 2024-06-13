using System;
using System.Xml.Linq;

namespace WorksheetGenerator.Utilities
{
    public static class HF
    {
        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static void Test()
        {
            Console.WriteLine("test");
        }

        public static XElement? GetElementOnly(XElement? element)
        {
            if (element != null)
            {
                // Create a new element with the same name and attributes
                XElement newElement = new XElement(element.Name);

                foreach (XAttribute attribute in element.Attributes())
                {
                    newElement.Add(new XAttribute(attribute));
                }

                return newElement;
            }
            else
            {
                return null;
            }
        }

        public static XElement? GetDocumentAndBodyOnly(XDocument doc)
        {
            XElement? originalDocElement = doc.Element(w + "document");
            XElement? docElement = GetElementOnly(doc.Element(w + "document"));
            if (docElement != null && originalDocElement != null)
            {
                XElement? body = GetElementOnly(originalDocElement.Element(w + "body"));
                if (body != null)
                {
                    docElement.Add(body);
                    return docElement;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public static bool IsIdentifier(XElement element)
        {
            XElement? runElement = element.Descendants(w + "r").FirstOrDefault();
            XElement? textElement = element.Descendants(w + "t").FirstOrDefault();
            if (textElement != null && runElement != null)
            {
                // Check if the text is all uppercase
                bool isUpperCase = textElement.Value.All(char.IsUpper);

                // Check if the text is bold
                XElement? runProperties = runElement.Element(w + "rPr");
                if (runProperties != null)
                {
                    bool isBold = runProperties.Element(w + "b") != null;
                    return isBold && isUpperCase;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static bool StartsWith(XElement element, string str)
        {
            return ((string)element).Trim().ToLower().StartsWith(str.ToLower());
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
                    {
                        break;
                    }

                    if (((string)paragraph).Trim().StartsWith(identifierName))
                    {
                        isBetweenIdentifiers = true;
                        result.Clear();
                    }
                }
                else if (isBetweenIdentifiers)
                {
                    if (!StartsWith(paragraph, "chatgpt:"))
                    {
                        result.Add(paragraph);
                    }
                }
            }

            return result;
        }

        public static bool IsImage(XElement element)
        {
            foreach (XElement child in element.Elements())
            {
                // Check if the child element is a drawing or picture
                if (child.Name == w + "drawing" || child.Name == w + "pict")
                    return true;

                // Check if the child element is a run element containing drawing or picture
                if (child.Name == w + "r")
                    foreach (XElement runChild in child.Elements())
                    {
                        if (runChild.Name == w + "drawing" || runChild.Name == w + "pict")
                            return true;
                    }
            }

            return false;
        }

        public static XElement? GetTitleElement(IEnumerable<XElement> elements)
        {
            foreach (XElement element in elements)
            {
                if (StartsWith(element, "title:"))
                    return element;
            }

            return null;
        }

        public static void AddTitleStyles(XElement element)
        {
            element.Add(new XElement(w + "pPr",
                new XElement(w + "jc",
                    new XAttribute(w + "val", "center")
                ),
                new XElement(w + "rPr",
                    new XElement(w + "b"),
                    new XElement(w + "bCs"),
                    new XElement(w + "sz",
                        new XAttribute(w + "val", "36")
                    ),
                    new XElement(w + "szCs",
                        new XAttribute(w + "val", "36")
                    ),
                    new XElement(w + "lang",
                        new XAttribute(w + "val", "en-US")
                    )
                )
            ));
        }

        public static List<XElement> GetProcessedReading(IEnumerable<XElement> allParagraphs)
        {
            List<XElement> paragraphs = GetParagraphsByIdentifier(allParagraphs, "READING");
            List<XElement> result = new List<XElement>();
            XElement? origTitleElement = GetTitleElement(paragraphs);

            // Format title
            if (origTitleElement != null)
            {
                XElement? newTitleElement = GetElementOnly(origTitleElement);
                if (newTitleElement != null)
                {
                    AddTitleStyles(newTitleElement);
                    newTitleElement.Value = origTitleElement.Value.ToUpper();
                    result.Add(newTitleElement);
                }
            }

            return result;
        }
    }
}