using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [XmlRoot("manifest", ElementName = "manifest", IsNullable = false)]
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

        [XmlElement("file-entry", ElementName = "file-entry")]
        public List<FileEntry> fileEntries { get; set; }

        public class FileEntry
        {
            [XmlAttribute("full-path", Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
            public string FullPath { get; set; } = "";

            [XmlAttribute("media-type", Namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")]
            public string MediaType { get; set; } = "";
        }
    }
}
