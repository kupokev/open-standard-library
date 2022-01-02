using System.Xml.Serialization;

namespace OoxSpreadsheet.Models.Files.xlsx
{
    [XmlRoot("Relationships", Namespace = "http://schemas.openxmlformats.org/package/2006/relationships")]
    public class Rels
    {
        public Rels()
        {
            Relationships = new List<Relationship>()
            {
                new Relationship() { Id = "rId3", Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", Target = "docProps/app.xml" },
                new Relationship() { Id = "rId2", Type = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", Target = "docProps/core.xml" },
                new Relationship() { Id = "rId1", Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target = "xl/workbook.xml" }
            };
        }

        [XmlElement(IsNullable = false, ElementName = "Relationships", Namespace = "http://schemas.openxmlformats.org/package/2006/relationships")]
        public List<Relationship> Relationships { get; set; }

        [XmlType("Relationship", Namespace = "http://schemas.openxmlformats.org/package/2006/relationships")]
        public class Relationship
        {
            [XmlAttribute]
            public string Id { get; set; }

            [XmlAttribute]
            public string Type { get; set; }

            [XmlAttribute]
            public string Target { get; set; }
        }
    }
}
