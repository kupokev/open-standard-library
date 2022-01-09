using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [XmlRoot("document-styles", ElementName = "document-styles", IsNullable = false)]
    public class ODStyles
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        public ODStyles()
        {
            xmlns = new XmlSerializerNamespaces();
            xmlns.Add("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
            xmlns.Add("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            xmlns.Add("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
            xmlns.Add("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
            xmlns.Add("draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0");
            xmlns.Add("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
            xmlns.Add("xlink", "http://www.w3.org/1999/xlink");
            xmlns.Add("dc", "http://purl.org/dc/elements/1.1/");
            xmlns.Add("number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0");
            xmlns.Add("svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0");
            xmlns.Add("of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2");

            fontFaceDecals = new();
            defaultStyles = new();
            automaticStyles = new();
            masterStyles = new();
        }

        [XmlAttribute(Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public string version { get; set; } = "1.3";

        [XmlElement("font-face-decls", ElementName = "font-face-decls")]
        public FontFaceDecals fontFaceDecals { get; set; }

        [XmlElement("styles", ElementName = "styles")]
        public DefaultStyles defaultStyles { get; set; }

        [XmlElement("automatic-styles", ElementName = "automatic-styles")]
        public AutomaticStyles automaticStyles { get; set; }

        [XmlElement("master-styles", ElementName = "master-styles")]
        public MasterStyles masterStyles { get; set; }


        public class FontFaceDecals
        {
            public FontFaceDecals()
            {
                fontFace = new() { new FontFace() };
            }

            [XmlElement("font-face", ElementName = "font-face")]
            public List<FontFace> fontFace { get; set; }

            public class FontFace
            {
                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string Name { get; set; } = "Calibri";

                [XmlAttribute("font-family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")]
                public string FontFamily { get; set; } = "Calibri";
            }
        }

        public class DefaultStyles
        {
            public DefaultStyles()
            {
                numberStyle = new();
                style = new();
                defaultStyle = new();
            }

            [XmlElement("number-style", ElementName = "number-style")]
            public NumberStyle numberStyle { get; set; }

            [XmlElement("style", ElementName = "style")]
            public Style style { get; set; }

            [XmlElement("default-style", ElementName = "default-style")]
            public DefaultStyle defaultStyle { get; set; }

            public class NumberStyle
            {
                public NumberStyle()
                {
                    Numbers = new() { new Number() };
                }

                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Name { get; set; } = "NO";

                [XmlElement("number", ElementName = "number")]
                public List<Number> Numbers { get; set; }

                public class Number
                {
                    [XmlAttribute("min-integer-digits", Namespace = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")] // namespace:number
                    public string MinimumIntegerDigits { get; set; } = "1";
                }
            }

            public class Style
            {
                public Style()
                {
                    tableCellProperties = new();
                    textProperties = new();
                }

                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Name { get; set; } = "Default";

                [XmlAttribute("family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Family { get; set; } = "table-cell";

                [XmlAttribute("data-style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string DataStyleName { get; set; } = "NO";

                [XmlElement("table-cell-properties", ElementName = "table-cell-properties")]
                public TableCellProperties tableCellProperties { get; set; }

                [XmlElement("text-properties", ElementName = "text-properties")]
                public TextProperties textProperties { get; set; }

                public class TableCellProperties
                {
                    [XmlAttribute("vertical-align", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string VerticalAlign { get; set; } = "automatic";

                    [XmlAttribute("background-color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string BackgroundColor { get; set; } = "transparent";
                }

                public class TextProperties
                {
                    [XmlAttribute("color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string Color { get; set; } = "#000000";

                    [XmlAttribute("font-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontName { get; set; } = "Calibri";

                    [XmlAttribute("font-name-asian", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontNameAsian { get; set; } = "Calibri";

                    [XmlAttribute("font-name-complex", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontNameComplex { get; set; } = "Calibri";

                    [XmlAttribute("font-size", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string FontSize { get; set; } = "11pt";

                    [XmlAttribute("font-size-asian", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontSizeAsian { get; set; } = "11pt";

                    [XmlAttribute("font-size-complex", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string FontSizeCopmlex { get; set; } = "11pt";
                }
            }

            public class DefaultStyle
            {
                public DefaultStyle()
                {
                    graphicProperties = new();
                }

                [XmlAttribute("family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                public string Family { get; set; } = "graphic";

                [XmlElement("graphic-properties", ElementName = "graphic-properties")]
                public GraphicProperties graphicProperties { get; set; }

                public class GraphicProperties
                {
                    [XmlAttribute("fill", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string Fill { get; set; } = "solid";

                    [XmlAttribute("fill-color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string FillColor { get; set; } = "#4472c4";

                    [XmlAttribute("opacity", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string Opacity { get; set; } = "100%";

                    [XmlAttribute("stroke", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string Stroke { get; set; } = "solid";

                    [XmlAttribute("stroke-width", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeWidth { get; set; } = "0.01389in";

                    [XmlAttribute("stroke-color", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeColor { get; set; } = "#2f528f";

                    [XmlAttribute("stroke-opacity", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeOpacity { get; set; } = "100%";

                    [XmlAttribute("stroke-linejoin", Namespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")] // namespace:draw
                    public string StrokeLineJoin { get; set; } = "miter";

                    [XmlAttribute("stroke-linecap", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")] // namespace:svg
                    public string StrokeLinecap { get; set; } = "butt";
                }
            }
        }

        public class AutomaticStyles
        {

        }

        public class MasterStyles
        {

        }
    }
}
