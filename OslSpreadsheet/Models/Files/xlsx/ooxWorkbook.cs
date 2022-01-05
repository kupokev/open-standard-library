using System.Xml;
using System.Xml.Serialization;

namespace OoxSpreadsheet.Models.Files.xlsx
{
    [XmlRoot("workbook", ElementName = "workbook", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class oWorkbook
    {
        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns;

        public oWorkbook()
        {
            fileVersion = new FileVersion();

            Sheets = new List<ooxSheet> { new ooxSheet(_sheetIndex) };

            xmlns = new XmlSerializerNamespaces();
            xmlns.Add("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            xmlns.Add("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        }

        [XmlIgnore]
        private int _sheetIndex
        {
            get
            {
                if (Sheets != null && sheets.Any())
                    return Sheets.Max(x => x.index) + 1;
                else
                    return 1;
            }
        }

        [XmlElement(IsNullable = false)]
        public FileVersion fileVersion { get; set; }

        [XmlArray(IsNullable = false)]
        public List<sheet> sheets
        {
            get
            {
                return Sheets.Select(x => new sheet()
                {
                    name = string.Format("Sheet{0}", x.index),
                    id = "rId" + x.index.ToString(),
                    sheetId = x.index.ToString()
                })
                    .ToList();
            }
        }

        public class FileVersion
        {
            [XmlAttribute]
            public string appName { get; set; } = "xl";

            [XmlAttribute]
            public string lastEdited { get; set; } = "7";

            [XmlAttribute]
            public string lowestEdited { get; set; } = "7";

            [XmlAttribute]
            public string rupBuild { get; set; } = "24701";

        }

        [XmlRoot("sheet", ElementName = "sheet")]
        public class sheet
        {

            [XmlAttribute]
            public string name { get; set; } = "";

            [XmlAttribute]
            public string sheetId { get; set; } = "";

            [XmlAttribute("id", Namespace = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships")]
            public string id { get; set; } = "";
        }

        public int AddSheet()
        {
            try
            {
                // Save id to return correct index
                var index = _sheetIndex;

                // Create new sheet
                Sheets.Add(new ooxSheet(index));

                // Return saved index
                return index;
            }
            catch
            {
                return -1;
            }
        }

        [XmlIgnore]
        public List<ooxSheet> Sheets { get; set; }
    }
}
