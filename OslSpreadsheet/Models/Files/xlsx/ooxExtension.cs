using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.xlsx
{
    [XmlRoot(ElementName = "Default", IsNullable = false)]
    public class ooxExtension
    {
        [XmlAttribute]
        public string Extension { get; set; }

        [XmlAttribute]
        public string ContentType { get; set; }
    }
}
