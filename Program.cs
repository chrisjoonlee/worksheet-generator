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
            IEnumerable<XElement> paragraphs = originalDoc.Descendants(HF.w + "p");

            // Create new document
            XDocument newDoc = new XDocument(
                HF.GetDocumentAndBodyOnly(originalDoc)
            );

            XElement? newBody = newDoc.Descendants(HF.w + "body").FirstOrDefault();

            foreach (XElement paragraph in HF.GetProcessedReading(paragraphs))
            {
                Console.WriteLine("Paragraph:", paragraph);
                newBody?.Add(paragraph);
            }

            Console.WriteLine(newBody);

            // newDoc.Save(filePath);
        }
    }
}