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

        public static XElement? ElementOnly(XElement? element)
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

        public static XElement? DocumentAndBodyOnly(XDocument doc)
        {
            XElement? originalDocElement = doc.Element(w + "document");
            XElement? docElement = ElementOnly(doc.Element(w + "document"));
            if (docElement != null && originalDocElement != null)
            {
                XElement? body = ElementOnly(originalDocElement.Element(w + "body"));
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
    }
}