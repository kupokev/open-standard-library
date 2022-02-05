using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.xlsx
{
    [XmlType("worksheet", TypeName = "worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class ooxSheet
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        // These are all ignored by XmlSerializer since they are not public properties
        private readonly SheetDimension _dimension;
        private readonly SheetFormatPr _sheetFormatPr;
        private readonly PageMargins _pageMargins;

        private readonly string _ignorable;
        private readonly string _uid;

        internal readonly int index;
        internal readonly string name;

        // This is needed for the XmlSerializer
        public ooxSheet() { }

        public ooxSheet(int id)
        {
            index = id;
            name = string.Format("Sheet{0}", id);

            _dimension = new SheetDimension();
            _sheetFormatPr = new SheetFormatPr();
            _pageMargins = new PageMargins();

            _ignorable = "x14ac xr xr2 xr3";
            _uid = "{" + Guid.NewGuid().ToString() + "}"; // I know there is a better way to do this

            SheetData = new List<SheetRow>();
            SheetViews = new List<SheetView> { new SheetView() };

            xmlns = new XmlSerializerNamespaces();
            xmlns.Add("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            xmlns.Add("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            xmlns.Add("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            xmlns.Add("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            xmlns.Add("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            xmlns.Add("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
        }

        // XmlAttributes

        [XmlAttribute(Namespace = "http://schemas.openxmlformats.org/markup-compatibility/2006")]
        public string Ignorable { get => _ignorable; set { } } // Leave set empty to make read only

        [XmlAttribute(Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")]
        public string uid { get => _uid; set { } } // Leave set empty to make read only


        // XmlElements

        [XmlElement(IsNullable = false, ElementName = "dimension", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public SheetDimension Dimension { get => _dimension; set { } }

        [XmlArray(IsNullable = false, ElementName = "sheetViews", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public List<SheetView> SheetViews { get; set; }

        [XmlElement(IsNullable = false, ElementName = "sheetFormatPr", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public SheetFormatPr SheetFormatPr { get => _sheetFormatPr; set { } }

        [XmlArray(IsNullable = false, ElementName = "sheetData", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public List<SheetRow> SheetData { get; set; }

        [XmlElement(IsNullable = false, ElementName = "pageMargins", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
        public PageMargins PageMargins { get => _pageMargins; set { } }
    }
}
