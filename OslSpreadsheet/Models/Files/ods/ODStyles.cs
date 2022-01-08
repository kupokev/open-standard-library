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

        }

        public class AutomaticStyles
        {

        }

        public class MasterStyles
        {

        }
    }
}
