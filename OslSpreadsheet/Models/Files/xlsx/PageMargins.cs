using System.Xml.Serialization;

namespace OoxSpreadsheet.Models.Files.xlsx
{
    [XmlType(TypeName = "pageMargins")]
    public class PageMargins
    {
        public PageMargins()
        {
            Left = "0.7";
            Right = "0.7";
            Top = "0.75";
            Bottom = "0.75";
            Header = "0.3";
            Footer = "0.3";
        }

        [XmlAttribute("left")]
        public string Left { get; set; }

        [XmlAttribute("right")]
        public string Right { get; set; }

        [XmlAttribute("top")]
        public string Top { get; set; }

        [XmlAttribute("bottom")]
        public string Bottom { get; set; }

        [XmlAttribute("header")]
        public string Header { get; set; }

        [XmlAttribute("footer")]
        public string Footer { get; set; }
    }
}
