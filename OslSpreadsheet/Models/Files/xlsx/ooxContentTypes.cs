using System.Xml.Serialization;

namespace OoxSpreadsheet.Models.Files.xlsx
{
    [XmlRoot("Types", ElementName = "Types", Namespace = "http://schemas.openxmlformats.org/package/2006/content-types", IsNullable = false)]
    public class ooxContentTypes
    {
        // Have to have a parameterless constructor to be able to serialize class
        public ooxContentTypes() { }

        public ooxContentTypes(oWorkbook workbook)
        {
            Extensions = new List<ooxExtension>();

            Extensions.Add(new ooxExtension() { Extension = "rels", ContentType = "application/vnd.openxmlformats-package.relationships+xml" });
            Extensions.Add(new ooxExtension() { Extension = "xml", ContentType = "application/xml" });

            Parts = new List<ooxPart>();

            Parts.Add(new ooxPart() { PartName = "/xl/workbook.xml", ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" });

            foreach (var sheet in workbook.sheets)
            {
                Parts.Add(new ooxPart() { PartName = string.Format("/xl/worksheets/{0}.xml", sheet.name), ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" });
            }

            Parts.Add(new ooxPart() { PartName = "/xl/theme/theme1.xml", ContentType = "application/vnd.openxmlformats-officedocument.theme+xml" });
            Parts.Add(new ooxPart() { PartName = "/xl/styles.xml", ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" });
            Parts.Add(new ooxPart() { PartName = "/docProps/core.xml", ContentType = "application/vnd.openxmlformats-package.core-properties+xml" });
            Parts.Add(new ooxPart() { PartName = "/docProps/app.xml", ContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml" });
        }

        [XmlElement(ElementName = "Default", IsNullable = false)]
        public List<ooxExtension> Extensions { get; set; }

        [XmlElement(ElementName = "Override", IsNullable = false)]
        public List<ooxPart> Parts { get; set; }
    }
}
