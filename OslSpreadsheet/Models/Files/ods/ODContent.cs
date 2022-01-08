using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [XmlRoot("document-content", ElementName = "document-content", IsNullable = false)]
    public class ODContent
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

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
        }

        [XmlAttribute(Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public string version { get; set; } = "1.3";
    }
}
