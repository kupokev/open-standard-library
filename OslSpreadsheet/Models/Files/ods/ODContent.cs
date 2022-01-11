using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [XmlRoot("document-content", ElementName = "document-content", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", IsNullable = false)]
    public class ODContent
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        private Body _body;

        public ODContent()
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
            automaticStyles = new();
            _body = new();
        }

        [XmlAttribute(Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public string version { get; set; } = "1.3";

        [XmlElement("font-face-decls", ElementName = "font-face-decls", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public FontFaceDecals fontFaceDecals { get; set; }

        [XmlElement("automatic-styles", ElementName = "automatic-styles", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public AutomaticStyles automaticStyles { get; set; }

        [XmlElement("body", ElementName = "body", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public Body body { get => _body; set { } }

        public class FontFaceDecals
        {
            public FontFaceDecals()
            {
                fontFace = new() { new FontFace() };
            }

            [XmlElement("font-face", ElementName = "font-face", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public List<FontFace> fontFace { get; set; }

            public class FontFace
            {
                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string Name { get; set; } = "Calibri";

                [XmlAttribute("font-family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")]
                public string FontFamily { get; set; } = "Calibri";
            }
        }

        public class AutomaticStyles
        {
            public AutomaticStyles()
            {
                automaticStyles = new List<Style>()
                {
                    new Style()
                    {
                        Name = "ce1",
                        Family = "table-cell",
                        DataStyleName = "N0"
                    },
                    new Style()
                    {
                        Name = "co1",
                        Family = "table-column",
                        tableColumnProperties = new ()
                    },
                    new Style()
                    {
                        Name = "ro1",
                        Family = "table-row",
                        tableRowProperties = new ()
                    },
                    new Style()
                    {
                        Name = "ta1",
                        Family = "table",
                        MasterPageName = "mp1",
                        tableProperties = new ()
                    }
                };
            }

            [XmlElement("style", ElementName = "style", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
            public List<Style> automaticStyles { get; set; }

            public class Style
            {
                [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? Name { get; set; }

                [XmlAttribute("family", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? Family { get; set; }

                [XmlAttribute("parent-style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? ParentStyleName { get; set; }

                [XmlAttribute("data-style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? DataStyleName { get; set; }

                [XmlAttribute("master-page-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")]
                public string? MasterPageName { get; set; }

                [XmlElement("table-column-properties", ElementName = "table-column-properties")]
                public TableColumnProperties tableColumnProperties { get; set; }

                [XmlElement("table-row-properties", ElementName = "table-row-properties")]
                public TableRowProperties tableRowProperties { get; set; }

                [XmlElement("table-properties", ElementName = "table-properties")]
                public TableProperties tableProperties { get; set; }

                public class TableColumnProperties
                {
                    [XmlAttribute("break-before", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string BreakBefore { get; set; } = "auto";

                    [XmlAttribute("column-width", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? ColumnWidth { get; set; } = "1.69333333333333cm";
                }

                public class TableRowProperties
                {
                    [XmlAttribute("row-height", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? RowHeight { get; set; } = "15pt";

                    [XmlAttribute("use-optimal-row-height", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? UseOptimalRowHeight { get; set; } = "true";

                    [XmlAttribute("break-before", Namespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")] // namespace:fo
                    public string BreakBefore { get; set; } = "auto";
                }

                public class TableProperties
                {
                    [XmlAttribute("display", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? Display { get; set; } = "true";

                    [XmlAttribute("writing-mode", Namespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0")] // namespace:style
                    public string? WritingMode { get; set; } = "lr-tb";
                }
            }
        }

        public class Body
        {
            public Body()
            {
                spreadsheet = new();
            }

            [XmlElement("spreadsheet", ElementName = "spreadsheet", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
            public Spreadsheet spreadsheet { get; set; }

            public class Spreadsheet
            {
                public Spreadsheet()
                {
                    calculationSettings = new();
                    Tables = new();

                    // For testing only
                    Tables.Add(new()
                    {
                        Name = "Sheet1"
                    });
                }

                [XmlElement("calculation-settings", ElementName = "calculation-settings", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public CalculationSettings calculationSettings { get; set; }

                [XmlElement("table", ElementName = "table", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public List<Table> Tables { get; set; }


                public class CalculationSettings
                {
                    [XmlAttribute("case-sensitive", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? CaseSensitive { get; set; } = "true";

                    [XmlAttribute("search-criteria-must-apply-to-whole-cell", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? SearchCriteriaMustApplyToWholeCell { get; set; } = "true";

                    [XmlAttribute("use-wildcards", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? UseWildcards { get; set; } = "false";

                    [XmlAttribute("use-regular-expressions", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? UseRegulareExpressions { get; set; } = "false";

                    [XmlAttribute("automatic-find-labels", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? AutomaticFindLabels { get; set; } = "false";
                }
            }
        }

        public class Table
        {
            private readonly int _maxColumns;
            private readonly int _maxRows;

            private List<TableRow> _rows;

            public Table()
            {
                _maxColumns = 16384;
                _maxRows = 1048577;

                tableColumn = new() { NumberColumnsRepeated = _maxColumns.ToString() };

                _rows = new()
                {
                    new TableRow()
                    {
                        RowIndex = 1,
                        NumberRowsRepeated = (_maxRows - 1).ToString(),
                        Cells = new()
                        {
                            new TableRow.TableCell()
                            {
                                NumberColumnsRepeated = _maxColumns.ToString(),
                                StyleName = null,
                                ValueType = null
                            }
                        }
                    }
                };
            }

            [XmlAttribute("name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
            public string Name { get; set; }

            [XmlAttribute("style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
            public string StyleName { get; set; } = "ta1";

            [XmlElement("table-column", ElementName = "table-column", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public TableColumn tableColumn { get; set; }

            [XmlElement("table-row", ElementName = "table-row", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
            public List<TableRow> Rows { get => _rows; set { } }

            public class TableColumn
            {
                [XmlAttribute("style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string StyleName { get; set; } = "co1";

                [XmlAttribute("number-columns-repeated", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string NumberColumnsRepeated { get; set; } = "";

                [XmlAttribute("default-cell-style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string DefaultCellStyleName { get; set; } = "ce1";
            }

            public class TableRow
            {
                public TableRow()
                {
                    Cells = new();
                }

                [XmlIgnore]
                public int RowIndex { get; set; }

                [XmlAttribute("number-rows-repeated", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string? NumberRowsRepeated { get; set; }

                [XmlAttribute("style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                public string StyleName { get; set; } = "ro1";

                [XmlElement("table-cell", ElementName = "table-cell", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")]
                public List<TableCell> Cells { get; set; }

                public class TableCell
                {
                    [XmlAttribute("value-type", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")] // namespace:office
                    public string? ValueType { get; set; } = "string";

                    [XmlAttribute("number-columns-repeated", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? NumberColumnsRepeated { get; set; }

                    [XmlAttribute("style-name", Namespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0")] // namespace:table
                    public string? StyleName { get; set; } = "ce1";
                }
            }
        }
    }
}
