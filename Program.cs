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
using DocumentFormat.OpenXml.ExtendedProperties;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;


namespace WorksheetGenerator
{
    class WorksheetGenerator
    {
        public static Styles Styles = new(
            El.Style(
                "Text",
                "Text",
                null,
                new ParagraphProperties(
                    new SpacingBetweenLines()
                    {
                        Line = "276",
                        LineRule = LineSpacingRuleValues.Auto,
                        Before = "0",
                        After = "0"
                    }

                ),
                new StyleRunProperties(
                    new Color() { Val = "000000", ThemeColor = ThemeColorValues.Text1, ThemeShade = "15" },
                    new RunFonts() { Ascii = "Aptos" },
                    new FontSize() { Val = "28" },
                    new FontSizeComplexScript() { Val = "48" }
                )
            ),
            El.Style(
                "Paragraph",
                "Paragraph",
                "Text",
                new ParagraphProperties(
                    new SpacingBetweenLines() { After = "280" }
                )
            ),
            El.Style(
                "IndentedText",
                "Indented Text",
                "Text",
                new ParagraphProperties(
                    new Indentation() { Left = "720" }
                )
            ),
            El.Style(
                "WorksheetTitle",
                "Worksheet Title",
                null,
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
                null,
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center },
                    new SpacingBetweenLines()
                    {
                        After = "400"
                    }
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
                "SubsectionTitle",
                "Subsection Title",
                "Text",
                new ParagraphProperties(
                    new SpacingBetweenLines()
                    {
                        After = "280"
                    }
                ),
                new StyleRunProperties(
                    new Bold(),
                    new BoldComplexScript()
                )
            ),
            El.Style(
                "AnswerKeyTitle",
                "Answer Key Title",
                null,
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                new StyleRunProperties(
                    new Bold(),
                    new BoldComplexScript(),
                    new Color() { Val = "000000", ThemeColor = ThemeColorValues.Text1, ThemeShade = "15" },
                    new RunFonts() { Ascii = "Aptos" },
                    new FontSize() { Val = "40" },
                    new FontSizeComplexScript() { Val = "40" }
                )
            ),
            El.Style(
                "AnswerKeySectionTitle",
                "Answer Key Section Title",
                "Text",
                null,
                new StyleRunProperties(
                    new Bold(),
                    new BoldComplexScript()
                )
            ),
            El.Style(
                "ListActivity",
                "List Activity",
                "Text",
                new ParagraphProperties(
                    new SpacingBetweenLines()
                    {
                        After = "280"
                    }
                )
            ),
            El.Style(
                "NoBorderTable",
                "No Border Table",
                null,
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
                null,
                new TableProperties(
                    new TableWidth()
                    {
                        Width = "5000",
                        Type = TableWidthUnitValues.Pct
                    },
                    El.TableBorders(BorderValues.Single, 24, ThemeColorValues.Accent4)
                )
            ),
            El.Style(
                "VocabBox",
                "Vocab Box",
                "Text",
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center },
                    new SpacingBetweenLines()
                    {
                        Line = "440",
                        LineRule = LineSpacingRuleValues.Auto,
                        Before = "0",
                        After = "220"
                    }
                ),
                new StyleRunProperties(
                    new Bold(),
                    new BoldComplexScript()
                )
            ),
            El.Style(
                "InlineImage",
                "Inline Image",
                null,
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center },
                    new SpacingBetweenLines()
                    {
                        Before = "240",
                        After = "400"
                    }
                )
            )
        );

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
                numberingPart.Numbering = new();

                // Styles
                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = Styles;
                styles.Save(stylePart);

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