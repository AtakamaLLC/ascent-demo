using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;

namespace ascent_demo
{
    class Program
    {

        public static void PopulateSpreadsheet(SpreadsheetDocument doc)
        {
            var wbpart = doc.AddWorkbookPart();
            wbpart.Workbook = new Workbook();

            var sheets = doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // sheets
            for (uint i = 1; i <= 10; i++)
            {
                var sheetData = new SheetData();

                var wspart = wbpart.AddNewPart<WorksheetPart>();
                wspart.Worksheet = new Worksheet(sheetData);

                var sheet = new Sheet()
                {
                    Id = wbpart.GetIdOfPart(wspart),
                    SheetId = i,
                    Name = $"sheet{i}"
                };
                sheets.Append(sheet);

                // rows
                for (uint j = 1; j < 100; j++)
                {
                    var row = new Row() { RowIndex = j };
                    sheetData.Append(row);

                    // cells
                    for (uint k = 1; k <= 200; k++)
                    {
                        var cell = new Cell
                        {
                            CellValue = new CellValue((int)(i * j * k)),
                            DataType = new EnumValue<CellValues>(CellValues.Number)
                        };
                        row.InsertAt(cell, 0);
                    }
                }
            }
        }

        public static void CreateSpreadsheet_MultiWrite(string path)
        {
            string fname = Path.Combine(path, Path.GetRandomFileName()) + ".xlsx";

            // given a path, OpenXml opens a file stream and writes to it repeatedly, resulting in multiple OS Write() calls
            var doc = SpreadsheetDocument.Create(fname, SpreadsheetDocumentType.Workbook);

            PopulateSpreadsheet(doc);
            doc.Close();
        }

        public static void CreateSpreadsheet_SingleWrite(string path)
        {
            var ms = new MemoryStream();

            // given a memory stream, OpenXml writes to memory repeatedly, not to disk
            var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);

            PopulateSpreadsheet(doc);
            doc.Close();

            string fname = Path.Combine(path, Path.GetRandomFileName()) + ".xlsx";

            // do a single Write() to disk when the document is ready
            File.WriteAllBytes(fname, ms.ToArray());
        }

        static void Main(string[] args)
        {
            try
            {
                string test_name = args[0];
                string path = args[1];

                Directory.CreateDirectory(path);

                if (test_name == "multi")
                {
                    CreateSpreadsheet_MultiWrite(path);
                }
                else if (test_name == "single")
                {
                    CreateSpreadsheet_SingleWrite(path);
                }
                else
                {
                    Console.WriteLine("first param should be 'multi' or 'single'");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception: {e}");
            }
        }
    }

}
