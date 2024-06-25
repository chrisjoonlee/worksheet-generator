using WorksheetGeneratorLibrary.Utilities;
using WorksheetGeneratorLibrary.Elements;
using WorksheetGeneratorLibrary.StyleList;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace WorksheetGenerator
{
    class WorksheetGenerator
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: dotnet run original.docx new.docx");
                return;
            }

            // Open file & create a new package
            string filePath = $"docs/{args[0]}";
            using (WordprocessingDocument origPackage = WordprocessingDocument.Open(filePath, false))
            using (WordprocessingDocument newPackage = WordprocessingDocument.Create($"docs/{args[1]}", WordprocessingDocumentType.Document))
            {
                if (origPackage is null)
                    throw new ArgumentNullException(nameof(origPackage));

                // Get original document contents
                MainDocumentPart origMainPart = origPackage.MainDocumentPart!;
                Body origBody = origMainPart!.Document!.Body!;
                OpenXmlElementList origElementList = origBody.ChildElements;

                // Populate new package
                (MainDocumentPart mainPart, Body body) = El.PopulateNewWordPackage(newPackage, 539);

                // Copy all images
                Dictionary<string, string> imageRelIds = []; // <old, new>
                foreach (ImagePart origImagePart in origMainPart.ImageParts)
                {
                    ImagePart newImagePart = mainPart.AddImagePart(origImagePart.ContentType);

                    // Copy image data
                    using Stream sourceStream = origImagePart.GetStream();
                    using Stream destStream = newImagePart.GetStream(FileMode.Create);
                    sourceStream.CopyTo(destStream);

                    // Keep track of image part relationship IDs
                    imageRelIds.Add(origMainPart.GetIdOfPart(origImagePart), mainPart.GetIdOfPart(newImagePart));

                    mainPart.Document.Save();
                }

                // Worksheet title
                Paragraph? worksheetTitleElement = HF.GetWorksheetTitleElement(origElementList);
                if (worksheetTitleElement != null)
                    body.AppendChild(worksheetTitleElement);

                // Keep track of section numbers
                int sectionNo = 1;

                // Vocab section
                (List<OpenXmlElement> vocabParagraphs, List<Paragraph> vocabAnswerKey) = HF.GetProcessedVocab(mainPart, origElementList, sectionNo);
                if (vocabParagraphs.Count > 0)
                {
                    sectionNo++;
                    body.Append(vocabParagraphs);
                }

                // Reading section
                List<Paragraph> readingParagraphs = HF.GetProcessedReading(origElementList, imageRelIds, sectionNo);
                if (readingParagraphs.Count > 0)
                {
                    sectionNo++;
                    body.Append(readingParagraphs);
                }

                // Comprehension questions section
                (List<OpenXmlElement> compQParagaphs, List<Paragraph> compQAnswerKey) = HF.GetProcessedCompQs(mainPart, origElementList, sectionNo);
                if (compQParagaphs.Count > 0)
                {
                    sectionNo++;
                    body.Append(compQParagaphs);
                }

                // Answer key
                body.Append(HF.AnswerKeyTitleElement());
                body.Append(vocabAnswerKey);
                body.Append(new Paragraph());
                body.Append(compQAnswerKey);
                body.Append(new Paragraph());

                origPackage.Dispose();
            }
        }
    }
}