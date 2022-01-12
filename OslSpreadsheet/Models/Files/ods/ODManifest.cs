using System.Xml;
using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
    [XmlRoot("manifest", ElementName = "manifest", Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0", IsNullable = false)]
    public class ODManifest
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        public ODManifest()
        {
            xmlns = new XmlSerializerNamespaces();
            xmlns.Add("manifest", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");

            fileEntries = new List<FileEntry>()
            {
                new FileEntry() { FullPath = "/", MediaType = "application/vnd.oasis.opendocument.spreadsheet" },
                new FileEntry() { FullPath = "styles.xml", MediaType = "text/xml" },
                new FileEntry() { FullPath = "content.xml", MediaType = "text/xml" },
                new FileEntry() { FullPath = "meta.xml", MediaType = "text/xml" }
            };
        }

        [XmlElement("file-entry", ElementName = "file-entry", Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
        public List<FileEntry> fileEntries { get; set; }


        [XmlType(AnonymousType = true, Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
        public class FileEntry
        {
            [XmlAttribute("full-path", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
            public string FullPath { get; set; } = "";

            [XmlAttribute("media-type", Form = System.Xml.Schema.XmlSchemaForm.Qualified, Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
            public string MediaType { get; set; } = "";
        }
    }
}
