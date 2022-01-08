namespace OslSpreadsheet.Models
{
    public class oWorkbook
    {
        public oWorkbook()
        {
            Sheets = new List<oSpreadsheet>();
        }

        private int _nextSheetIndex
        {
            get => Sheets.Any() ? Sheets.Max(x => x.Index) + 1 : 1;
        }

        /// <summary>
        /// Application that generated this file
        /// </summary>
        public string Generator { get; set; } = "Open Standard Library v1";

        public string InitialCreator { get; set; } = "";

        public string Creator { get; set; } = "";

        public string CreationDate { get; set; } = "";

        public List<oSpreadsheet> Sheets { get; set; }

        public oSpreadsheet AddSheet()
        {
            var index = _nextSheetIndex;

            return _AddSheet(new oSpreadsheet(index, string.Format("Sheet{0}", index)));
        }

        public oSpreadsheet AddSheet(string name)
        {
            var index = _nextSheetIndex;

            return _AddSheet(new oSpreadsheet(index, name));
        }

        public async Task<oSpreadsheet> AddSheetAsync()
        {
            var index = _nextSheetIndex;

            return await Task.Run(() => _AddSheet(new oSpreadsheet(index, string.Format("Sheet{0}", index))));
        }

        public async Task<oSpreadsheet> AddSheetAsync(string name)
        {
            var index = _nextSheetIndex;

            return await Task.Run(() => _AddSheet(new oSpreadsheet(index, name)));
        }

        private oSpreadsheet _AddSheet(oSpreadsheet sheet)
        {
            try
            {
                lock (this)
                {
                    var e = Sheets.FirstOrDefault(x => x.Index == sheet.Index);

                    if (e == null)
                        Sheets.Add(sheet);
                    else
                        e = sheet;

                    return e ?? sheet;
                }
            }
            catch
            {
                throw new Exception("There was an error addind a new worksheet to the workbook");
            }
        }
    }
}
