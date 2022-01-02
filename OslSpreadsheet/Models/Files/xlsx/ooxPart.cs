using System.Xml.Serialization;

namespace OoxSpreadsheet.Models.Files.xlsx
{
    [XmlRoot(ElementName = "Override", IsNullable = false)]
    public class ooxPart
    {
        [XmlAttribute]
        public string PartName { get; set; }

        [XmlAttribute]
        public string ContentType { get; set; }
    }
}
