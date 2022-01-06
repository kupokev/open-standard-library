namespace OslSpreadsheet.Models
{
    public class oSpreadsheet
    {
        private readonly int _index;

        private readonly List<oCell> _cells;

        public oSpreadsheet(int index, string name)
        {
            _index = index;

            _cells = new();

            SheetName = name;
        }

        public string SheetName { get; set; }

        public List<oCell> Cells { get => _cells; }

        public int Index { get => _index; }

        public oCell AddCell(int row, int column)
        {
            return _AddCell(new oCell(row, column));
        }

        public async Task<oCell> AddCellAsync(int row, int column)
        {
            return await Task.Run(() => _AddCell(new oCell(row, column)));
        }

        private oCell _AddCell(oCell cell)
        {
            try
            {
                lock (this)
                {
                    var e = Cells.FirstOrDefault(x => x.Row == cell.Row && x.Column == cell.Column);

                    if (e == null)
                        Cells.Add(cell);
                    else
                        e = cell;

                    return e ?? cell;
                }
            }
            catch
            {
                throw new Exception("There was an error addind a new worksheet to the workbook");
            }
        }
    }
}
