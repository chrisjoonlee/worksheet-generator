using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorksheetGeneratorLibrary.Excel;
using WorksheetGeneratorLibrary.Utilities;
using WorksheetGeneratorLibrary.PowerPoint;
using WorksheetGeneratorLibrary.Word;
using WXML = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Presentation;

namespace CIExcelToWord
{
    class CIExcelToWord
    {
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: dotnet run original.xlsx new.docx language");
                return;
            }

            // Paths
            string excelFilePath = $"docs/{args[0]}";
            string baseFileName = Path.GetFileNameWithoutExtension(args[0]);
            string wordFilePath = $"docs/{args[1]}";
            string pptFilePath = $"docs/{baseFileName}_new.pptx";
            string pptntFilePath = $"docs/{baseFileName}_new_no_text.pptx";
            string imagesFolderPath = $"docs/{baseFileName}-imgs";

            // Language
            string language = args[2].ToLower();

            // Open Excel file, create Word package
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
            using (WordprocessingDocument newPackage = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
            {
                if (spreadsheetDocument is null)
                    throw new ArgumentNullException(nameof(spreadsheetDocument));

                // Get excel parts
                WorkbookPart? workbookPart = spreadsheetDocument.WorkbookPart;
                if (workbookPart is null)
                    throw new ArgumentNullException(nameof(workbookPart));
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStringTable is null)
                    throw new ArgumentNullException(nameof(sharedStringTable));
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Access the media folder and extract images
                List<string> imageFilePaths = EF.ExtractImages(excelFilePath, imagesFolderPath);

                // Populate new Word package
                (MainDocumentPart mainPart, WXML.Body body) = WF.PopulateNewWordPackage(newPackage, 1134, "blue");
                // WF.AddPageNumbers(mainPart);

                // Create PowerPoint packages
                PresentationDocument pptDoc = PresentationDocument.Create(pptFilePath, PresentationDocumentType.Presentation);
                PresentationPart pptPresentationPart = pptDoc.AddPresentationPart();
                pptPresentationPart.Presentation = new Presentation();
                PF.CreatePresentationParts(pptPresentationPart);

                PresentationDocument pptntDoc = PresentationDocument.Create(pptntFilePath, PresentationDocumentType.Presentation);
                PresentationPart pptntPresentationPart = pptntDoc.AddPresentationPart();
                pptntPresentationPart.Presentation = new Presentation();
                PF.CreatePresentationParts(pptntPresentationPart);

                // Get excel data
                List<List<Cell>> rows = EF.GetRows(sheetData);

                // Establish which columns to read from
                int mainColIndex = 0;
                int imageColIndex = 0;
                int choiceColIndex = 0;

                List<Cell> headerRow = rows[0];
                for (int i = 0; i < headerRow.Count; i++)
                {
                    Cell cell = headerRow[i];
                    if (EF.IsTextCell(cell))
                    {
                        string text = EF.GetCellText(cell, sharedStringTable);
                        if (text.ToLower().StartsWith("image"))
                            imageColIndex = i;
                        if (text.ToLower().StartsWith(language))
                        {
                            mainColIndex = i;
                            choiceColIndex = i + 1;
                        }
                    }
                }

                List<List<List<Cell>>> sections = HF.GetExcelSections(rows.Skip(1).ToList(), imageColIndex, sharedStringTable);

                // MAIN SECTION
                List<List<Cell>> mainSection = sections[0];

                // Get chapter # and title
                string? chapterNo = "";
                string title = "";
                int titleRowIndex = 0;
                foreach (List<Cell> cells in mainSection)
                {
                    // Find image cell that contains a number
                    Cell imageCell = cells[imageColIndex];
                    if (EF.IsNumberCell(imageCell))
                    {
                        chapterNo = EF.GetNumberAsString(imageCell);

                        // Get title
                        Cell mainCell = cells[mainColIndex];
                        if (EF.IsTextCell(mainCell))
                            title = EF.GetCellText(mainCell, sharedStringTable);
                        break;
                    }

                    titleRowIndex++;
                }
                if (string.IsNullOrWhiteSpace(chapterNo))
                    throw new NullReferenceException("No chapter number provided");
                if (string.IsNullOrWhiteSpace(title))
                    throw new NullReferenceException("No title provided");

                // Add chapter # and title to Word doc
                WF.AppendToBody(body, WF.Paragraph($"CHAPTER {chapterNo}", "ChapterTitle"));
                WF.AppendToBody(body, WF.Paragraph(title, "ChapterSubtitle"));
                WF.AppendToBody(body, WF.Paragraph());
                WF.AppendToBody(body, WF.SectionBreak(1134, "blue", 1));

                // Read rest of main section
                for (int i = titleRowIndex + 1; i < mainSection.Count; i++)
                {
                    List<Cell> cells = mainSection[i];

                    List<WXML.Paragraph> paragraphs = HF.GetImageAndTextFromExcel(
                        cells, mainPart, body,
                        imageColIndex, mainColIndex, choiceColIndex,
                        imagesFolderPath, sharedStringTable,
                        1440000
                    );

                    WF.AppendToBody(body, paragraphs);
                }
                WF.AppendToBody(body, WF.SectionBreak(1134, "blue", 2));
                WF.AppendToBody(body, WF.PageBreak());

                // ALL OTHER SECTIONS
                for (int i = 1; i < sections.Count; i++)
                {
                    List<List<Cell>> currentSection = sections[i];

                    // Parse header row
                    List<Cell> sectionHeaderRow = currentSection[0];
                    string sectionType = EF.GetCellText(sectionHeaderRow[imageColIndex], sharedStringTable).ToLower();
                    Console.WriteLine(sectionType);

                    // Summary section
                    if (sectionType.StartsWith("summary") || sectionType.StartsWith("review"))
                    {
                        List<WXML.Paragraph> paragraphs = HF.GetProcessedSummaryFromExcel(
                            currentSection, mainPart, body,
                            imageColIndex, mainColIndex, choiceColIndex,
                            imagesFolderPath, sharedStringTable,
                            1440000
                        );

                        WF.AppendToBody(body, paragraphs);
                    }

                    // Matching section
                    if (sectionType.StartsWith("match"))
                    {
                        List<OpenXmlElement> elements = HF.GetProcessedMatchingFromExcel(
                            currentSection, mainPart, body,
                            imageColIndex, mainColIndex, choiceColIndex,
                            imagesFolderPath, sharedStringTable,
                            1080000
                        );

                        WF.AppendToBody(body, elements);
                    }

                    // True or false section
                    if (sectionType.StartsWith("true") || sectionType.StartsWith("t/f") || sectionType.StartsWith("t / f"))
                    {
                        List<OpenXmlElement> elements = HF.GetProcessedTrueOrFalseFromExcel(
                            currentSection, mainPart, body,
                            imageColIndex, mainColIndex, choiceColIndex,
                            imagesFolderPath, sharedStringTable,
                            1080000
                        );

                        WF.AppendToBody(body, elements);
                    }
                }

                newPackage.Dispose();
            }
        }
    }
}