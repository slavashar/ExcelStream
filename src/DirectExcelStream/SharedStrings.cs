using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml;

namespace ExcelStream
{
    public class SharedStrings
    {
        public static readonly Uri Url = new Uri("/xl/sharedStrings.xml", UriKind.Relative);

        private readonly IDictionary<string, int> sharedStrings = new Dictionary<string, int>();

        public int GetSharedStringIndex(string text)
        {
            int index;
            if (!this.sharedStrings.TryGetValue(text, out index))
            {
                this.sharedStrings.Add(text, index = this.sharedStrings.Count);
            }

            return index;
        }
        
        public PackagePart CreatePart(Package package)
        {
            var part = package.CreatePart(
                Url, 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", 
                CompressionOption.Normal);

            using (var writer = XmlWriter.Create(part.GetStream()))
            {
                this.WriteStringTable(writer);
            }

            return part;
        }

        private void WriteStringTable(XmlWriter writer)
        {
            writer.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("count", this.sharedStrings.Count.ToString());
            writer.WriteAttributeString("uniqueCount", this.sharedStrings.Count.ToString());

            foreach (var str in this.sharedStrings.OrderBy(x => x.Value).Select(x => x.Key))
            {
                writer.WriteStartElement("si");

                writer.WriteStartElement("t");

                if (str.StartsWith(" ") || str.EndsWith(" "))
                {
                    writer.WriteAttributeString("xml", "space", null, "preserve");
                }

                writer.WriteString(str);
                writer.WriteEndElement();

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }
    }
}
