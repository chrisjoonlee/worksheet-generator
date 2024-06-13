using System;
using System.Xml.Linq;
using WorksheetGenerator.Utilities;

namespace WorksshetGenerator
{
    class WorksheetGenerator
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: dotnet run <folder_name>");
                return;
            }

            // Load document
            string filePath = $"docs/{args[0]}/word/document.xml";
            XDocument originalDoc = XDocument.Load(filePath);
            XDocument newDoc = new XDocument(
                HF.GetDocumentAndBodyOnly(originalDoc)
            );

            // Get all paragraphs
            IEnumerable<XElement> paragraphs = originalDoc.Descendants(HF.w + "p");

            // Process reading
            HF.ProcessReading(paragraphs);
        }
    }
}