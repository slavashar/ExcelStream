using System;
using System.Globalization;
using System.Xml;

namespace ExcelStream
{
    public class Row : IDisposable
    {
        private readonly XmlWriter worksheetWriter;
        private readonly int rowIndex;
        private readonly Func<string, int> getSharedStringIndex;

        private int lastColumnIndex = -1;

        public Row(XmlWriter worksheetWriter, int rowIndex, Func<string, int> getSharedStringIndex)
        {
            this.worksheetWriter = worksheetWriter;
            this.rowIndex = rowIndex;
            this.getSharedStringIndex = getSharedStringIndex;

            this.worksheetWriter.WriteStartElement("row");
        }

        public void Dispose()
        {
            // close row
            this.worksheetWriter.WriteEndElement();
        }

        public string WriteEmpty(string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                return cell.Ref;
            }
        }

        public string WriteInlineString(string value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteAttributeString("t", "inlineStr");

                this.worksheetWriter.WriteStartElement("is");
                this.worksheetWriter.WriteElementString("t", value);
                this.worksheetWriter.WriteEndElement();

                return cell.Ref;
            }
        }

        public string WriteValue(string value, string cellReference = null, Style style = null)
        {
            if (value == null)
            {
                throw new ArgumentException("value is not provided", value);
            }

            using (var cell = new CellWriter(this, cellReference, style))
            {
                int index = this.getSharedStringIndex(value.ToString());

                this.worksheetWriter.WriteAttributeString("t", "s");
                this.worksheetWriter.WriteElementString("v", index.ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(long value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteElementString("v", value.ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(int value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteElementString("v", value.ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(decimal value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteElementString("v", value.ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(double value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteElementString("v", value.ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(DateTime value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteElementString("v", TicksToOADate(value.Ticks).ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(TimeSpan value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteElementString("v", value.TotalDays.ToString());

                return cell.Ref;
            }
        }

        public string WriteValue(bool value, string cellReference = null, Style style = null)
        {
            using (var cell = new CellWriter(this, cellReference, style))
            {
                this.worksheetWriter.WriteAttributeString("t", "b");
                this.worksheetWriter.WriteElementString("v", value ? "1" : "0");

                return cell.Ref;
            }
        }

        public string WriteValue(object value, string cellReference = null, Style style = null)
        {
            if (value == null || value as string == string.Empty)
            {
                return this.WriteEmpty(cellReference, style);
            }
            else if (value is byte || value is short || value is int || value is long || value is ushort || value is uint || value is ulong)
            {
                using (var cell = new CellWriter(this, cellReference, style))
                {
                    this.worksheetWriter.WriteElementString("v", value.ToString());

                    return cell.Ref;
                }
            }
            else if (value is decimal || value is float || value is double)
            {
                using (var cell = new CellWriter(this, cellReference, style))
                {
                    this.worksheetWriter.WriteElementString("v", value.ToString());

                    return cell.Ref;
                }
            }
            else if (value is DateTime)
            {
                return this.WriteValue((DateTime)value, cellReference, style);
            }
            else if (value is TimeSpan)
            {
                return this.WriteValue((TimeSpan)value, cellReference, style);
            }
            else if (value is string)
            {
                return this.WriteValue((string)value, cellReference, style);
            }
            else if (value is bool)
            {
                return this.WriteValue((bool)value, cellReference, style);
            }
            else
            {
                throw new InvalidOperationException();
            }
        }

        private static int ParseColumnIndex(string columnName)
        {
            var first = (char)columnName[0] - (int)'A';

            if (!char.IsLetter(columnName[1]))
            {
                return first;
            }

            var second = (char)columnName[1] - (int)'A';

            return ((first + 1) * 26) + second;
        }

        private static string ColumnIndexToName(int columnIndex)
        {
            var second = (char)(((int)'A') + columnIndex % 26);

            columnIndex /= 26;

            if (columnIndex == 0)
                return second.ToString();
            else
                return new string(new char[] { (char)(((int)'A') - 1 + columnIndex), second });
        }

        private static string RowIndexToName(int rowIndex)
        {
            return (rowIndex + 1).ToString();
        }

        private static double TicksToOADate(long value)
        {
            if (value == 0)
            {
                return 0.0;
            }

            long millis = (value - 599264352000000000) / 10000;
            if (millis < 0)
            {
                long frac = millis % 864000000000;
                if (frac != 0) millis -= (864000000000 + frac) * 2;
            }

            return (double)millis / 864000000000;
        }

        private class CellWriter : IDisposable
        {
            private readonly Row row;
            private readonly string cellRef;

            public CellWriter(Row row, string cellReference = null, Style style = null)
            {
                this.row = row;

                this.row.worksheetWriter.WriteStartElement("c");

                int columnIndex;
                if (cellReference != null)
                {
                    columnIndex = ParseColumnIndex(cellReference);

                    if (columnIndex != this.row.lastColumnIndex + 1)
                    {
                        this.row.worksheetWriter.WriteAttributeString("r", cellReference);
                    }

                    this.row.lastColumnIndex = columnIndex;
                }
                else
                {
                    columnIndex = ++this.row.lastColumnIndex;
                }

                if (style != null)
                {
                    this.row.worksheetWriter.WriteAttributeString("s", style.Id.ToString());
                }

                this.cellRef = cellReference ?? ColumnIndexToName(columnIndex) + RowIndexToName(this.row.rowIndex);
            }

            public string Ref
            {
                get { return this.cellRef; }
            }

            public void Dispose()
            {
                this.row.worksheetWriter.WriteEndElement();
            }
        }
    }
}
