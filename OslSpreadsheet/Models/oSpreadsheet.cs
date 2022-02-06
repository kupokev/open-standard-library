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

        public oCell AddCell(int row, int column, string value)
        {
            return _AddCell(new oCell(row, column)
            {
                Value = value
            });
        }

        public oCell AddCell(int row, int column, string value, CellValueType valueType)
        {
            return _AddCell(new oCell(row, column)
            {
                Value = value,
                ValueType = valueType
            });
        }

        public async Task<oCell> AddCellAsync(int row, int column)
        {
            return await Task.Run(() => AddCell(row, column));
        }

        public async Task<oCell> AddCellAsync(int row, int column, string value)
        {
            return await Task.Run(() => AddCell(row, column, value));
        }

        public async Task<oCell> AddCellAsync(int row, int column, string value, CellValueType valueType)
        {
            return await Task.Run(() => AddCell(row, column, value, valueType));
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

        /// <summary>
        /// Get a count of columns in a spreadsheet
        /// </summary>
        public int ColumnCount { get => _cells.Max(x => x.Column); }

        /// <summary>
        /// Get a count of rows in the spreadsheet
        /// </summary>
        public int RowCount { get => _cells.Max(x => x.Row); }

        /// <summary>
        /// Get all cells from a particular row
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public List<oCell> GetRow(int index)
        {
            return _cells.Where(x => x.Row == index).OrderBy(x => x.Column).ToList();
        }

        /// <summary>
        /// Converts a spreadsheet to a 2D array
        /// </summary>
        /// <returns></returns>
        public string[,] ToArray()
        {
            var retVal = new string[RowCount, ColumnCount];

            foreach(var c in _cells)
            {
                retVal[c.Row - 1, c.Column - 1] = c.Value;
            }

            return retVal;
        }
    }
}
