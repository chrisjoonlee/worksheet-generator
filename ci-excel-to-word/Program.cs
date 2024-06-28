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
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: dotnet run original.xlsx new.docx");
                return;
            }

            // Paths
            string excelFilePath = $"docs/{args[0]}";
            string baseFileName = Path.GetFileNameWithoutExtension(args[0]);
            string wordFilePath = $"docs/{args[1]}";
            string imagesFolderPath = $"docs/{baseFileName}-imgs";

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
                List<string> imageFilePaths = Excel.ExtractImages(excelFilePath, imagesFolderPath);

                // Populate new Word package
                (MainDocumentPart mainPart, WXML.Body body) = El.PopulateNewWordPackage(newPackage);

                // Read excel text
                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (Excel.IsTextCell(cell))
                        {
                            // Get text
                            string text = Excel.GetCellText(cell, sharedStringTable);
                            Console.WriteLine(text);

                            // Place in document
                            body.AppendChild(
                                new WXML.Paragraph(new WXML.Run(new WXML.Text(text)))
                            );
                        }
                        else if (Excel.IsImageCell(cell))
                        {
                            // Get image path
                            string? imagePath = Excel.GetImagePath(cell, imagesFolderPath);
                            if (imagePath != null)
                                Console.WriteLine(imagePath);
                        }
                    }
                }

                newPackage.Dispose();
            }
        }
    }
}