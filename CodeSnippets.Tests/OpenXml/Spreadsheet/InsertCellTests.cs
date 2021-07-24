//
// InsertCellTests.cs
//
// Copyright 2020 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Spreadsheet
{
    public class InsertCellTests
    {
        [Fact]
        public void CanInsertCell()
        {
            using var stream = new MemoryStream();
            using (var spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                // Create an empty workbook.
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets());

                // Create an empty worksheet and add the worksheet to the workbook.
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                workbookPart.Workbook.Sheets.AppendChild(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    Name = "Test",
                    SheetId = 1
                });

                // This example correctly inserts a cell with an inline string,
                // noting that Excel always inserts shared strings as shown in
                // the next example.
                InsertCellWithInlineString(worksheetPart.Worksheet, 1, "C", "1C");

                // This example inserts a cell with a shared string that is
                // contained in the SharedStringTablePart. Note that the cell
                // value is the zero-based index of the SharedStringItem
                // contained in the SharedStringTable.
                var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTablePart.SharedStringTable =
                    new SharedStringTable(
                        new SharedStringItem(
                            new Text("2C")));

                InsertCellWithSharedString(worksheetPart.Worksheet, 2, "C", 0);
            }

            File.WriteAllBytes("WorkbookWithNewCells.xlsx", stream.ToArray());
        }

        private static void InsertCellWithInlineString(
            Worksheet worksheet,
            uint rowIndex,
            string columnName,
            string value)
        {
            InsertCell(worksheet, rowIndex, new Cell
            {
                CellReference = columnName + rowIndex,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value)),
            });
        }

        private static void InsertCellWithSharedString(
            Worksheet worksheet,
            uint rowIndex,
            string columnName,
            uint value)

        {
            InsertCell(worksheet, rowIndex, new Cell
            {
                CellReference = columnName + rowIndex,
                DataType = CellValues.SharedString,
                CellValue = new CellValue(value.ToString())
            });
        }

        private static void InsertCell(Worksheet worksheet, uint rowIndex, Cell cell)
        {
            SheetData sheetData = worksheet.Elements<SheetData>().Single();

            // Get or create a Row with the given rowIndex.
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };

                // The sample assumes that the newRow can simply be appended,
                // e.g., because rows are added in ascending order only.
                sheetData.AppendChild(row);
            }

            // The sample assumes two things: First, no cell with the same cell
            // reference exists. Second, cells are added in ascending order.
            // If that is not the case, you need to deal with that situation.
            row.AppendChild(cell);
        }
    }
}
