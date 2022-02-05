using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.xlsx
{
    [XmlType(TypeName = "sheetView", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class SheetView
    {
        public SheetView()
        {
            TabSelected = "1";
            WorkbookViewId = "0";
        }

        [XmlAttribute("tabSelected")]
        public string TabSelected { get; set; }

        [XmlAttribute("workbookViewId")]
        public string WorkbookViewId { get; set; }
    }
}
