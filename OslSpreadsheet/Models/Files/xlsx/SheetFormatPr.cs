using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.xlsx
{
    [XmlRoot("sheetFormatPr", ElementName = "sheetFormatPr", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class SheetFormatPr
    {
        public SheetFormatPr()
        {
            DefaultRowHeight = "15";
            DyDescent = "0.25";
        }

        [XmlAttribute("defaultRowHeight")]
        public string DefaultRowHeight { get; set; }

        [XmlAttribute("dyDescent", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")]
        public string DyDescent { get; set; }
    }
}
