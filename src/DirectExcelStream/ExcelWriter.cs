using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Xml;

namespace ExcelStream
{
    public class ExcelWriter : IDisposable
    {
        private readonly Package package;
        private readonly IList<Sheet> sheets = new List<Sheet>();

        private readonly Styles styles = new Styles();
        private readonly SharedStrings sharedStrings = new SharedStrings();

        public ExcelWriter(string filename)
        {
            this.package = Package.Open(filename, FileMode.Create);
        }

        public ExcelWriter(Stream stream)
        {
            this.package = Package.Open(stream, FileMode.Create);
        }

        public void Dispose()
        {
            using (this.package)
            {
                this.sharedStrings.CreatePart(this.package);

                this.styles.CreatePart(this.package);

                var workbookUrl = new Uri("/xl/workbook.xml", UriKind.Relative);

                var workbookPart = this.package
                    .CreatePart(workbookUrl, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    CompressionOption.Fast);

                using (var writer = XmlWriter.Create(workbookPart.GetStream()))
                {
                    this.WriteWorkbook(writer);
                }

                foreach (var sheet in this.sheets)
                {
                    workbookPart.CreateRelationship(
                        sheet.Url, 
                        TargetMode.Internal, 
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        "rId" + sheet.Id);
                }

                workbookPart.CreateRelationship(
                    SharedStrings.Url,
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");

                workbookPart.CreateRelationship(
                    Styles.Url,
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");

                this.package.CreateRelationship(workbookUrl,
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            }
        }

        public Font CreateFont()
        {
            return this.styles.CreateFont();
        }
        public Font CreateFont(bool bold = false, bool underline = false, string color = null)
        {
            var font = this.styles.CreateFont();
            font.Bold = bold;
            font.Underline = underline;
            font.Color = color;
            return font;
        }

        public Fill CreateFill()
        {
            return this.styles.CreateFill();
        }

        public Style CreateStyle()
        {
            return this.styles.CreateStyle();
        }

        public Style CreateStyle(HorizontalAlignment alignment = HorizontalAlignment.None, Fill fill = null, Font font = null, StyleFormat? numberFormat = null)
        {
            var style = this.styles.CreateStyle();
            style.HorizontalAlignment = alignment;
            style.Font = font;
            style.Fill = fill;
            style.NumberFormat = numberFormat;
            return style;
        }

        public Worksheet CreateWorksheet(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException("name");
            }

            var sh = new Sheet(sheets.Count + 1, name);

            this.sheets.Add(sh);

            var part = this.package.CreatePart(sh.Url, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Maximum);

            return new Worksheet(XmlWriter.Create(part.GetStream()), this.sharedStrings.GetSharedStringIndex);
        }

        public Worksheet InsertWorksheet(string name, int index)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentNullException("name");
            }

            var sh = new Sheet(sheets.Count + 1, name);

            this.sheets.Insert(index, sh);

            var part = this.package.CreatePart(sh.Url, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Maximum);
            
            return new Worksheet(XmlWriter.Create(part.GetStream()), this.sharedStrings.GetSharedStringIndex);
        }

        private void WriteWorkbook(XmlWriter writer)
        {
            writer.WriteStartElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            writer.WriteStartElement("sheets");

            foreach (var sheet in this.sheets)
            {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", sheet.Name);
                writer.WriteAttributeString("sheetId", sheet.Id.ToString());
                writer.WriteAttributeString("r", "id", null, "rId" + sheet.Id);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}
