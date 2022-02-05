using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.xlsx
{
    [XmlType(TypeName = "dimension")]
    public class SheetDimension
    {
        [XmlAttribute("ref")]
        public string reference { get; set; } = "A1";
    }
}
