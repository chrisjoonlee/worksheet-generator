using System;
using System.Xml.Linq;
using WorksheetGenerator.Utilities;
using WorksheetGenerator.Elements;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;


namespace WorksshetGenerator
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

                // Create document structure in new package
                MainDocumentPart mainPart = newPackage.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Add section properties & set page margins
                SectionProperties sectionProperties = new();
                PageMargin pageMargin = new();
                pageMargin.Top = 539;
                pageMargin.Right = 539;
                pageMargin.Bottom = 539;
                pageMargin.Left = 539;
                sectionProperties.Append(pageMargin);
                body.Append(sectionProperties);

                // Define styles
                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = new(
                    El.Style(
                        "WorksheetTitle",
                        "Worksheet Title",
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new StyleRunProperties(
                            new Bold(),
                            new BoldComplexScript(),
                            new Color() { Val = "000000", ThemeColor = ThemeColorValues.Text1, ThemeShade = "15" },
                            new RunFonts() { Ascii = "Aptos" },
                            new FontSize() { Val = "48" },
                            new FontSizeComplexScript() { Val = "48" }
                        )
                    )
                );
                styles.Save(stylePart);

                // Worksheet title
                Paragraph? worksheetTitleElement = HF.GetWorksheetTitleElement(origElementList);
                if (worksheetTitleElement != null)
                    body.AppendChild(worksheetTitleElement);

                // Keep track of section numbers
                int sectionNo = 1;

                // Vocab section
                (OpenXmlElementList vocabParagraphs, OpenXmlElementList vocabAnswerKey) = HF.GetProcessedVocab(origElementList, sectionNo);
                // if (vocabParagraphs.Count > 0)
                // {
                //     sectionNo++;
                //     foreach (OpenXmlElement paragraph in vocabParagraphs)
                //         body.Append(paragraph);
                // }

                // origPackage.Dispose();
            }


            //     // Reading section
            //     List<XElement> readingParagraphs = HF.GetProcessedReading(paragraphs, sectionNo);
            //     if (readingParagraphs.Count > 0)
            //     {
            //         sectionNo++;
            //         foreach (XElement paragraph in readingParagraphs)
            //             newBody.Add(paragraph);
            //     }

            //     // Comprehension questions section
            //     (List<XElement> compQParagaphs, List<XElement> compQAnswerKey) = HF.GetProcessedCompQs(paragraphs, sectionNo);

            //     // Answer key
            //     newBody.Add(HF.AnswerKeyTitleElement());
            //     newBody.Add(vocabAnswerKey);

            //     newDoc.Save(filePath);
            // }
        }
    }
}