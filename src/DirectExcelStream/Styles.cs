using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Xml;

namespace ExcelStream
{
    public class Styles
    {
        public static readonly Uri Url = new Uri("/xl/styles.xml", UriKind.Relative);

        private readonly IList<Font> fonts = new List<Font>();
        private readonly IList<Fill> fills = new List<Fill>();
        private readonly IList<Style> styles = new List<Style>();

        public Styles()
        {
        }

        public Font CreateFont()
        {
            var font = new Font(this.fonts.Count + 1);
            this.fonts.Add(font);
            return font;
        }

        public Fill CreateFill()
        {
            var fill = new Fill(this.fills.Count + 2);
            this.fills.Add(fill);
            return fill;
        }

        public Style CreateStyle()
        {
            var style = new Style(this.styles.Count + 1);
            this.styles.Add(style);
            return style;
        }

        public PackagePart CreatePart(Package package)
        {
            var part = package.CreatePart(
                Url, 
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", 
                CompressionOption.Normal);

            using (var writer = XmlWriter.Create(part.GetStream()))
            {
                this.WriteStyles(writer);
            }

            return part;
        }

        private void WriteStyles(XmlWriter writer)
        {
            writer.WriteStartElement("styleSheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            // fonts
            writer.WriteStartElement("fonts");
            writer.WriteAttributeString("count", (this.fonts.Count + 1).ToString());

            //default font
            writer.WriteStartElement("font");
            writer.WriteEndElement();

            foreach (var font in this.fonts)
            {
                writer.WriteStartElement("font");

                if (font.Bold)
                {
                    writer.WriteStartElement("b");
                    writer.WriteEndElement();
                }

                if (font.Underline)
                {
                    writer.WriteStartElement("u");
                    writer.WriteEndElement();
                }

                if (font.Color != null)
                {
                    writer.WriteStartElement("color");
                    writer.WriteAttributeString("rgb", font.Color);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            // fills
            writer.WriteStartElement("fills");
            writer.WriteAttributeString("count", (this.fills.Count + 2).ToString());

            //default fill
            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "none");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "gray125");
            writer.WriteEndElement();
            writer.WriteEndElement();

            foreach (var fill in this.fills)
            {
                writer.WriteStartElement("fill");
                writer.WriteStartElement("patternFill");
                writer.WriteAttributeString("patternType", "solid");

                if (fill.ForegroundColourIndexed.HasValue || fill.ForegroundColourRGB != null)
                {
                    writer.WriteStartElement("fgColor");

                    if (fill.ForegroundColourIndexed.HasValue)
                    {
                        writer.WriteAttributeString("indexed", fill.ForegroundColourIndexed.ToString());
                    }

                    if (fill.ForegroundColourRGB != null)
                    {
                        writer.WriteAttributeString("rgb", fill.ForegroundColourRGB);
                    }

                    writer.WriteEndElement();
                }

                if (fill.BackgroundColourIndexed.HasValue || fill.BackgroundColourRGB != null)
                {
                    writer.WriteStartElement("bgColor");

                    if (fill.BackgroundColourIndexed.HasValue)
                    {
                        writer.WriteAttributeString("indexed", fill.BackgroundColourIndexed.ToString());
                    }

                    if (fill.BackgroundColourRGB != null)
                    {
                        writer.WriteAttributeString("rgb", fill.BackgroundColourRGB);
                    }

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            // borders
            writer.WriteStartElement("borders");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("border");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // cellStyleXfs
            writer.WriteStartElement("cellStyleXfs");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("xf");
            writer.WriteEndElement();
            writer.WriteEndElement();

            // cellXfs
            writer.WriteStartElement("cellXfs");
            writer.WriteAttributeString("count", (this.styles.Count + 1).ToString());

            // default style
            writer.WriteStartElement("xf");
            writer.WriteEndElement();

            foreach (var style in this.styles)
            {
                writer.WriteStartElement("xf");

                if (style.NumberFormat.HasValue)
                {
                    writer.WriteAttributeString("numFmtId", ((int)style.NumberFormat).ToString());
                }

                if (style.Font != null)
                {
                    writer.WriteAttributeString("fontId", style.Font.Id.ToString());
                }

                if (style.Fill != null)
                {
                    writer.WriteAttributeString("fillId", style.Fill.Id.ToString());
                }

                if (style.NumberFormat.HasValue)
                {
                    writer.WriteAttributeString("applyNumberFormat", "1");
                }

                if (style.Font != null)
                {
                    writer.WriteAttributeString("applyFont", "1");
                }

                if (style.Fill != null)
                {
                    writer.WriteAttributeString("applyFill", "1");
                }

                if (style.HorizontalAlignment != HorizontalAlignment.None)
                {
                    writer.WriteAttributeString("applyAlignment", "1");
                }

                if (style.HorizontalAlignment != HorizontalAlignment.None)
                {
                    writer.WriteStartElement("alignment");
                    writer.WriteAttributeString("horizontal", style.HorizontalAlignment.ToString().ToLowerInvariant());
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            writer.WriteEndElement();
        }
    }
}
