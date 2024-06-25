using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WorksheetGeneratorLibrary.Excel;

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
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                // Read data using SAX syntax
                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        reader.ReadFirstChild();
                        do
                        {
                            if (reader.ElementType == typeof(Cell))
                            {
                                Cell cell = (Cell)reader.LoadCurrentElement();
                                string cellValue = Excel.GetCellValue(cell, sharedStringTable);
                                Console.Write(cellValue + " ");
                            }
                        } while (reader.ReadNextSibling());
                        Console.WriteLine();
                    }

                    // if (reader.ElementType == typeof(CellValue))
                    // {
                    //     text = reader.GetText();
                    //     Console.Write(text + " ");
                    // }
                }
            }
        }
    }
}