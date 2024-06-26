using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorksheetGeneratorLibrary.Excel;
using WorksheetGeneratorLibrary.Utilities;
using WorksheetGeneratorLibrary.Elements;
using DocumentFormat.OpenXml.Wordprocessing;
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

            string filePath = $"docs/{args[0]}";
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            using (WordprocessingDocument newPackage = WordprocessingDocument.Create($"docs/{args[1]}", WordprocessingDocumentType.Document))
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
                // DrawingsPart? drawingsPart = worksheetPart.DrawingsPart;
                // if (drawingsPart is null)
                //     throw new ArgumentNullException(nameof(drawingsPart));

                // Populate new package
                (MainDocumentPart mainPart, Body body) = El.PopulateNewWordPackage(newPackage);

                // Read excel data
                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (Excel.IsTextCell(cell))
                        {
                            string text = Excel.GetCellText(cell, sharedStringTable);
                            // Console.WriteLine(text);
                        }
                        else
                        {
                            // Console.WriteLine("Not text");
                        }
                    }
                }

                // Access the media folder and extract images
                using (FileStream zipToOpen = new FileStream(filePath, FileMode.Open))
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                {
                    var mediaEntries = archive.Entries.Where(e => e.FullName.StartsWith("xl/media/")).ToList();

                    if (!mediaEntries.Any())
                    {
                        Console.WriteLine("No images found in the workbook.");
                    }
                    else
                    {
                        foreach (var mediaEntry in mediaEntries)
                        {
                            // using (Stream stream = mediaEntry.Open())
                            // using (MemoryStream memoryStream = new MemoryStream())
                            // {
                            //     stream.CopyTo(memoryStream);
                            //     byte[] imageBytes = memoryStream.ToArray();

                            //     // Save the image or process it as needed
                            //     string newImagePath = $"docs/{mediaEntry.Name}";
                            //     File.WriteAllBytes(newImagePath, imageBytes);
                            //     Console.WriteLine($"Image saved to {newImagePath}");
                            // }
                            Console.WriteLine(mediaEntry);
                        }
                    }
                }
            }
        }
    }
}