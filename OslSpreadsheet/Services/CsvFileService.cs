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
        public Task<oWorkbook> GenerateModel(byte[] file)
        {
            throw new NotImplementedException();
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
