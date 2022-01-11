using System.Xml.Serialization;

namespace OslSpreadsheet.Models.Files.ods
{
    [XmlRoot("document-meta", ElementName = "document-meta", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", IsNullable = false)]
    public class ODMeta
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        public ODMeta()
        {
            xmlns = new XmlSerializerNamespaces();
            xmlns.Add("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
            xmlns.Add("meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0");
            xmlns.Add("dc", "http://purl.org/dc/elements/1.1/");
            xmlns.Add("xlink", "http://www.w3.org/1999/xlink");

            meta = new();
        }

        [XmlAttribute(Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public string version { get; set; } = "1.3";

        [XmlElement("meta", ElementName = "meta", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
        public Meta meta { get; set; }

        public class Meta
        {
            private static readonly string _appName = "Open Standard Library v1.0.0";
            private readonly DateTime _date;

            private DateTime? _creationDate;

            public Meta()
            {
                _date = DateTime.Now;
            }

            [XmlElement("generator", ElementName = "generator", Namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0")]
            public string Generator { get; set; } = _appName;

            //[XmlElement("initial-creator", ElementName = "initial-creator")]
            //public string InitialCreator { get; set; } = _appName;

            [XmlElement("creator", ElementName = "creator", Namespace = "http://purl.org/dc/elements/1.1/")]
            public string Creator { get; set; } = _appName;

            [XmlElement("creation-date", ElementName = "creation-date", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
            public string CreationDate 
            { 
                get
                {
                    return _creationDate == null ? DateTime.Now.ToString("u").Replace(' ', 'T') : ((DateTime)_creationDate).ToString("u").Replace(' ', 'T');
                }
                set
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(value))
                            _creationDate = DateTime.Now;
                        else
                            _creationDate = DateTime.Parse(value.Replace('T', ' '));
                    }
                    catch { }
                }
            }

            [XmlElement("date", ElementName = "date", Namespace = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")]
            public string Date { get => _date.ToString("u").Replace(' ', 'T'); set { } }
        }
    }
}
