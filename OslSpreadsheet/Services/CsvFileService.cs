using OslSpreadsheet.Models;
using System.Text;

namespace OslSpreadsheet.Services
{
    internal class CsvFileService : IFileService
    {
        /// <summary>
        /// Converts an oWorkbook object to a CSV file byte array.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            if (workbook.Sheets.Count == 0) throw new ArgumentException("Workbook requires at least 1 sheet to convert to CSV");

            byte[] output;

            string result = "";

            var cols = workbook.Sheets[0].Cells.Max(x => x.Column);
            var rows = workbook.Sheets[0].Cells.Max(x => x.Row);

            var cells = workbook.Sheets[0].Cells;

            // Have to start at 1 instead of 0
            for (int r = 1; r <= rows; r++)
            {
                var values = new List<string>();

                for (int c = 1; c <= cols; c++)
                {
                    string value = cells.FirstOrDefault(x => x.Row == r && x.Column == c)?.Value ?? "";
                    values.Add(ParseValue(value));
                }

                result += "\"" + string.Join("\",\"", values) + "\"" + Environment.NewLine;
            }

            output = Encoding.UTF8.GetBytes(result);

            return Task.FromResult(output);
        }
         
        /// <summary>
        /// Manipulates values of a cell to proper values that can be stored in CSV
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private string ParseValue(string value)
        {
            // Double quotes in CSV are escaped by adding another double quote in front of it according to RFC-4180
            return value.Replace("\"", "\"\"");
        }

        /// <summary>
        /// Converts a CSV file byte array to an oWorkbook object
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public async Task<oWorkbook> GenerateModel(byte[] file)
        {
            // Create new workbook
            oWorkbook workbook = new();

            // Add sheet to workbook
            var sheet1 = await workbook.AddSheetAsync();

            try
            {
                // Check 
                if (file == null || file.Length == 0) throw new ArgumentNullException("File is empty");

                // Convert file to string and replace line endings with Environment.NewLine
                var contents = Encoding.UTF8.GetString(file).ReplaceLineEndings(Environment.NewLine);

                // Split the contents into multiple rows
                var lines = contents.Split(Environment.NewLine).ToList();

                // Remove blank lines
                lines.Remove("");

                for (int r = 0; r < lines.Count(); r++)
                {
                    // Split the line to columns
                    var cols = lines[r].Replace("\", \"", "\",\"").Replace("\" ,\"", "\",\"").Split("\",\"");

                    // Remove the first double-quote
                    cols[0] = cols[0].Remove(0, 1);

                    // Remove end quote for the last element
                    cols[cols.Count() - 1] = cols[cols.Count() - 1].Remove(cols[cols.Count() - 1].Length - 1, 1);

                    // Cycle through the columns and add to list
                    for (int c = 0; c < cols.Count(); c++)
                    {
                        sheet1.AddCell(r + 1, c + 1, cols[c]);
                    }
                }
            }
            catch
            {

            }

            return workbook;
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public ValueTask DisposeAsync()
        {
            throw new NotImplementedException();
        }
    }
}
