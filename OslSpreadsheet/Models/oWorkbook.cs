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

        public List<oSpreadsheet> Sheets { get; set; }

        public oSpreadsheet AddSheet()
        {
            var index = _nextSheetIndex;

            return _AddSheet(new oSpreadsheet(index, String.Format("Sheet{0}", index)));
        }

        public oSpreadsheet AddSheet(string name)
        {
            var index = _nextSheetIndex;

            return _AddSheet(new oSpreadsheet(index, name));
        }

        public async Task<oSpreadsheet> AddSheetAsync()
        {
            var index = _nextSheetIndex;

            return await Task.Run(() => _AddSheet(new oSpreadsheet(index, String.Format("Sheet{0}", index))));
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
