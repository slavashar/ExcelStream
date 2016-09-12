using System;
using System.Collections.Generic;
using System.Xml;

namespace ExcelStream
{
    public class Worksheet : IDisposable
    {
        private readonly XmlWriter worksheetWriter;
        private readonly Func<string, int> getSharedStringIndex;
        private readonly IList<Hyperlink> hyperLinks;

        private int lastRowIndex = -1;

        public Worksheet(XmlWriter worksheetWriter, Func<string, int> getSharedStringIndex)
        {
            this.worksheetWriter = worksheetWriter;
            this.getSharedStringIndex = getSharedStringIndex;

            this.worksheetWriter.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            this.worksheetWriter.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            this.worksheetWriter.WriteStartElement("sheetData");

            //Initialise hyperlinks collection
            this.hyperLinks = new List<Hyperlink>();
        }

        public void Dispose()
        {
            using (this.worksheetWriter)
            {
                // close sheetData
                this.worksheetWriter.WriteEndElement();

                this.WriteHyperlinks();

                // close worksheet
                this.worksheetWriter.WriteEndElement();
            }
        }

        private void WriteHyperlinks()
        {
            if (hyperLinks.Count > 0)
            {
                this.worksheetWriter.WriteStartElement("hyperlinks");

                foreach (var hyperlink in this.hyperLinks)
                {
                    this.worksheetWriter.WriteStartElement("hyperlink");
                    this.worksheetWriter.WriteAttributeString("ref", hyperlink.Ref);
                    this.worksheetWriter.WriteAttributeString("location", hyperlink.Location);

                    if (hyperlink.Display != null)
                    {
                        this.worksheetWriter.WriteAttributeString("display", hyperlink.Display);
                    }

                    this.worksheetWriter.WriteEndElement();
                }

                this.worksheetWriter.WriteEndElement();
            }
        }

        public Row CreateRow()
        {
            return new Row(worksheetWriter, ++this.lastRowIndex, getSharedStringIndex);
        }

        public void AddHyperlink(string @ref, string location, string display = null)
        {
            if (@ref == null)
            {
                throw new ArgumentNullException("ref");
            }

            if (location == null)
            {
                throw new ArgumentNullException("location");
            }

            this.hyperLinks.Add(new Hyperlink(@ref, location, display));
        }
    }
}
