using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using OpenXmlSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace ExcelStream.Facts
{
    public class ExcelWriterFacts
    {
        [Fact]
        public void crate_a_worksheet()
        {
            using (var stream = new System.IO.MemoryStream())
            {
                using (var writer = new ExcelWriter(stream))
                {
                    using (writer.CreateWorksheet("test"))
                    { }
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false))
                {
                    var test = spreadsheet.WorkbookPart.Workbook.Sheets.Cast<OpenXmlSheet>().Single();

                    Assert.Equal("test", test.Name);
                }
            }
        }

        [Fact]
        public void write_in_line_text()
        {
            using (var stream = new System.IO.MemoryStream())
            {
                using (var writer = new ExcelWriter(stream))
                {
                    using (var worksheet = writer.CreateWorksheet("test"))
                    {
                        using (var row = worksheet.CreateRow())
                        {
                            row.WriteInlineString("my text");
                        }
                    }
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false))
                {
                    var test = spreadsheet.WorkbookPart.Workbook.Sheets.Cast<OpenXmlSheet>().Single();

                    var worksheet = ((WorksheetPart)spreadsheet.WorkbookPart.GetPartById(test.Id)).Worksheet;

                    Assert.Equal("my text", worksheet.FirstChild.InnerText);
                }
            }
        }

        [Fact]
        public void write_text()
        {
            using (var stream = new System.IO.MemoryStream())
            {
                using (var writer = new ExcelWriter(stream))
                {
                    using (var worksheet = writer.CreateWorksheet("test"))
                    {
                        using (var row = worksheet.CreateRow())
                        {
                            row.WriteValue("my text");
                        }
                    }
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false))
                {
                    var test = spreadsheet.WorkbookPart.Workbook.Sheets.Cast<OpenXmlSheet>().Single();

                    var worksheet = ((WorksheetPart)spreadsheet.WorkbookPart.GetPartById(test.Id)).Worksheet;

                    Assert.Equal("0", worksheet.FirstChild.InnerText);
                }
            }
        }

        [Fact]
        public void write_number()
        {
            using (var stream = new System.IO.MemoryStream())
            {
                using (var writer = new ExcelWriter(stream))
                {
                    using (var worksheet = writer.CreateWorksheet("test"))
                    {
                        using (var row = worksheet.CreateRow())
                        {
                            row.WriteValue(100);
                        }
                    }
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false))
                {
                    var test = spreadsheet.WorkbookPart.Workbook.Sheets.Cast<OpenXmlSheet>().Single();

                    var worksheet = ((WorksheetPart)spreadsheet.WorkbookPart.GetPartById(test.Id)).Worksheet;

                    Assert.Equal("100", worksheet.FirstChild.InnerText);
                }
            }
        }
    }
}
