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
                PageMargin pageMargin = new()
                {
                    Top = 539,
                    Right = 539,
                    Bottom = 539,
                    Left = 539
                };
                sectionProperties.Append(pageMargin);
                body.Append(sectionProperties);

                // Numbering definitions
                NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart");
                numberingPart.Numbering = new(
                    new AbstractNum(
                        new Level(
                            new NumberingFormat() { Val = NumberFormatValues.Decimal },
                            new LevelText() { Val = "%1." },
                            new StartNumberingValue() { Val = 1 }
                        )
                    )
                    { AbstractNumberId = 1 },
                    new NumberingInstance(
                        new AbstractNumId() { Val = 1 }
                    )
                    { NumberID = 1 }
                );


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
                    ),
                    El.Style(
                        "SectionTitle",
                        "Section Title",
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center }
                        ),
                        new StyleRunProperties(
                            new Bold(),
                            new BoldComplexScript(),
                            new Color() { Val = "0F9ED5", ThemeColor = ThemeColorValues.Accent4 },
                            new RunFonts() { Ascii = "Aptos" },
                            new FontSize() { Val = "36" },
                            new FontSizeComplexScript() { Val = "36" }
                        )
                    ),
                    El.Style(
                        "NoBorderTable",
                        "No Border Table",
                        null,
                        null,
                        new TableProperties(
                            new TableWidth()
                            {
                                Width = "5000",
                                Type = TableWidthUnitValues.Pct
                            },
                            El.TableBorders(BorderValues.Nil, 0, ThemeColorValues.Background1)
                        )
                    ),
                    El.Style(
                        "Box",
                        "Box",
                        null,
                        null,
                        new TableProperties(
                            new TableWidth()
                            {
                                Width = "5000",
                                Type = TableWidthUnitValues.Pct
                            },
                            El.TableBorders(BorderValues.Single, 24, ThemeColorValues.Accent4)
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
                (List<OpenXmlElement> vocabParagraphs, List<OpenXmlElement> vocabAnswerKey) = HF.GetProcessedVocab(origElementList, sectionNo);
                if (vocabParagraphs.Count > 0)
                {
                    sectionNo++;
                    foreach (OpenXmlElement paragraph in vocabParagraphs)
                        body.Append(paragraph);
                }

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