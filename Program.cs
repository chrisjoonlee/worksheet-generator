using System;
using System.Xml.Linq;
using WorksheetGenerator.Utilities;
using WorksheetGenerator.Elements;

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
            IEnumerable<XElement> paragraphs = originalDoc.Descendants(El.w + "p");

            // Create new document
            XDocument newDoc = new XDocument(
                HF.GetDocumentAndBodyOnly(originalDoc)
            );

            // Get body
            XElement? newBody = newDoc.Descendants(El.w + "body").FirstOrDefault();
            if (newBody != null)
            {
                // Copy section properties
                XElement? sectionPr = originalDoc.Descendants(El.w + "sectPr").FirstOrDefault();
                if (sectionPr != null)
                    newBody.Add(sectionPr);

                // Set page margins
                El.SetPageMargins(newBody, 539);

                // Worksheet title
                XElement? worksheetTitleElement = HF.GetWorksheetTitleElement(paragraphs);
                if (worksheetTitleElement != null)
                    newBody.Add(worksheetTitleElement);

                // Keep track of section numbers
                int sectionNo = 1;

                // Vocab section
                (List<XElement> vocabParagraphs, List<XElement> vocabAnswerKey) = HF.GetProcessedVocab(paragraphs, sectionNo);
                if (vocabParagraphs.Count > 0)
                {
                    sectionNo++;
                    foreach (XElement paragraph in vocabParagraphs)
                        newBody.Add(paragraph);
                }

                // Reading section
                List<XElement> readingParagraphs = HF.GetProcessedReading(paragraphs, sectionNo);
                if (readingParagraphs.Count > 0)
                {
                    sectionNo++;
                    foreach (XElement paragraph in readingParagraphs)
                        newBody.Add(paragraph);
                }

                // Answer key
                newBody.Add(HF.AnswerKeyTitleElement());
                newBody.Add(vocabAnswerKey);

                newDoc.Save(filePath);
            }
        }
    }
}