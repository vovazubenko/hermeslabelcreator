using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GemBox.Spreadsheet;
using HermesLabelCreator.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace HermesLabelCreator.Helpers
{
    public static class ExcelManager
    {
        public static string ConvertCsvToExcel(string csvFileName)
        {
            string directoryPath = Path.GetDirectoryName(csvFileName);
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            // Create new CSV file.
            var csvFile = ExcelFile.Load(csvFileName, new CsvLoadOptions(CsvType.SemicolonDelimited));
            csvFile.Worksheets.Add("csv");

            string fileName = Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(csvFileName) + DateTime.Now.ToString("YYYYmmddhhmmss") + ".xlsx");

            csvFile.Save(fileName, new XlsxSaveOptions());
            
            return fileName;
        }

        public static string ConvertExcelToCsv(string excelfileFileName)
        {
            string directoryPath = Path.GetDirectoryName(excelfileFileName);
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            // Create new CSV file.
            var excelFile = ExcelFile.Load(excelfileFileName, new XlsxLoadOptions());
            excelFile.Save(Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(excelfileFileName) + ".csv"), new CsvSaveOptions(CsvType.SemicolonDelimited));

            return Path.Combine(directoryPath, Path.GetFileNameWithoutExtension(excelfileFileName) + ".csv");
        }

        public static SpreadsheetCellDto[][] GetRows(Stream xlsxFile)
        {
            MemoryStream documentStream = new MemoryStream();
            xlsxFile.CopyTo(documentStream);

            List<SpreadsheetCellDto[]> data = new List<SpreadsheetCellDto[]>();
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(documentStream, false))
                {
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                    SharedStringItem[] sharedStringItemsArray = workbookPart.SharedStringTablePart?.SharedStringTable
                        .Elements<SharedStringItem>()
                        .ToArray();
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    foreach (Row r in sheetData.Elements<Row>())
                    {
                        if (r.Elements<Cell>().Any(ce => ce.DataType != null))
                        {
                            List<SpreadsheetCellDto> row = new List<SpreadsheetCellDto>();
                            foreach (Cell c in r.Elements<Cell>())
                            {
                                string text = GetValueFromCell(c, sharedStringItemsArray);
                                row.Add(new SpreadsheetCellDto
                                {
                                    CellColumnName = Regex.Replace(c.CellReference.Value, @"[\d-]", string.Empty),
                                    Value = text
                                });
                            }

                            data.Add(row.ToArray());
                        }
                    }
                }
            }
            catch (FileFormatException)
            {
                throw new ArgumentException("Invalid spreadsheet file.");
            }

            return data.ToArray();
        }

        public static string GetValueFromCell(Cell cell, SharedStringItem[] sharedStringItems)
        {
            int id;
            string cellValue = cell.InnerText;

            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        id = int.Parse(cellValue);
                        SharedStringItem item = sharedStringItems[id];

                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }

                        break;
                }
            }

            return cellValue;
        }

        public static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        public static void UpdateCell(string text, string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            if (worksheetPart != null)
            {
                Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);

                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);

                worksheetPart.Worksheet.Save();
            }

        }

        private static Cell GetCell(Worksheet worksheet,
                  string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            string cellReference = columnName + rowIndex;
            string[] ar = row.Elements<Cell>().Select(c => c.CellReference.Value).ToArray();

            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }

        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.Elements<SheetData>().First().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }
}
