using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorksheetGeneratorLibrary.Excel;
using WorksheetGeneratorLibrary.Utilities;
using WorksheetGeneratorLibrary.Word;
using WXML = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
using System.Linq;
using System.IO;
using System.IO.Compression;

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

                // Get excel data
                IEnumerable<Row> rows = sheetData.Elements<Row>();

                // Establish which columns to read from
                int mainColIndex = 0;
                int imageColIndex = 0;
                int choiceColIndex = 0;

                List<Cell> headerRow = EF.GetCellList(rows.First());
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

                // Get chapter # and title
                string? chapterNo = "";
                string title = "";
                int titleRowIndex = 1;
                foreach (Row row in rows.Skip(1))
                {
                    List<Cell> cells = EF.GetCellList(row);

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
                WF.AppendToBody(body, WF.SectionBreak("blue"));

                // Read rest of excel sheet
                foreach (Row row in sheetData.Elements<Row>().Skip(titleRowIndex + 1))
                {
                    List<Cell> cells = EF.GetCellList(row);

                    // Image
                    Cell imageCell = cells[imageColIndex];
                    if (EF.IsImageCell(imageCell))
                        WF.AppendToBody(body, WF.Paragraph("IMAGE"));

                    // Main text
                    if (cells.Count > mainColIndex)
                    {
                        Cell mainCell = cells[mainColIndex];
                        if (EF.IsTextCell(mainCell))
                        {
                            string[] lines = EF.GetCellText(mainCell, sharedStringTable).Split('\n');
                            foreach (string line in lines)
                                WF.AppendToBody(body, WF.Paragraph(line, "TextCentered"));
                        }
                    }

                    // foreach (Cell cell in row.Elements<Cell>())
                    // {
                    //     if (EF.IsTextCell(cell))
                    //     {
                    //         // Get text
                    //         string text = EF.GetCellText(cell, sharedStringTable);
                    //         Console.WriteLine(text);

                    //         // Place in document
                    //         WF.AppendToBody(body, WF.Paragraph(text));
                    //     }
                    //     else if (EF.IsImageCell(cell))
                    //     {
                    //         // Get image path
                    //         string? imagePath = EF.GetImagePath(cell, imagesFolderPath);
                    //         if (imagePath != null)
                    //             Console.WriteLine(imagePath);
                    //     }
                    //     else if (EF.IsNumberCell(cell))
                    //     {
                    //         // Get number as string
                    //         Console.WriteLine(EF.GetNumberAsString(cell));
                    //     }
                    // }
                }

                WF.AppendToBody(body, WF.SectionBreak("blue", 2));

                newPackage.Dispose();
            }
        }
    }
}