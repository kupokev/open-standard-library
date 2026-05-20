using OslSpreadsheet.Models;
using OslSpreadsheet.Services;
using System.Text;

namespace OoxSpreadsheet
{
    public interface ISpreadsheet : IDisposable, IAsyncDisposable
    {
        oWorkbook Workbook { get; }

        /// <summary>
        /// Headers from the most recent ReadCsvRowsAsync call when hasHeaderRow was true.
        /// </summary>
        string[]? CsvHeaders { get; }

        Task<byte[]> GenerateCsvFileAsync();

        Task<byte[]> GenerateOdsFileAsync();

        Task<byte[]> GenerateXlsxFileAsync();

        Task<oWorkbook> ImportCsvFileAsync(byte[] file);

        Task<oWorkbook> ImportOdsFileAsync(byte[] file);

        Task<oWorkbook> ImportXlsxFileAsync(byte[] file);

        /// <summary>
        /// Reads CSV rows one at a time from a stream without loading the entire file into memory.
        /// When hasHeaderRow is true, the first row is consumed as headers (available via CsvHeaders) and not yielded.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="hasHeaderRow"></param>
        /// <param name="rowLimit"></param>
        /// <returns></returns>
        IAsyncEnumerable<string[]> ReadCsvRowsAsync(Stream stream, bool hasHeaderRow = false, int? rowLimit = null);
    }

    public class Spreadsheet : ISpreadsheet
    {
        private oWorkbook _workbook;

        public Spreadsheet()
        {
            _workbook = new oWorkbook();
        }

        public oWorkbook Workbook  { get => _workbook; }

        /// <inheritdoc />
        public string[]? CsvHeaders { get; private set; }


        public async Task<byte[]> GenerateCsvFileAsync()
        {
            IFileService _fileService = new DelimitedFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        public async Task<byte[]> GenerateOdsFileAsync()
        {
            IFileService _fileService = new OdsFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        public async Task<byte[]> GenerateXlsxFileAsync()
        {
            IFileService _fileService = new XlsxFileService();

            return await _fileService.GenerateFileAsync(Workbook);
        }

        public async Task<oWorkbook> ImportCsvFileAsync(byte[] file)
        {
            IFileService _fileService = new DelimitedFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        public async Task<oWorkbook> ImportOdsFileAsync(byte[] file)
        {
            IFileService _fileService = new OdsFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        public async Task<oWorkbook> ImportXlsxFileAsync(byte[] file)
        {
            IFileService _fileService = new XlsxFileService();

            _workbook = await _fileService.GenerateModel(file);

            return _workbook;
        }

        /// <inheritdoc />
        public async IAsyncEnumerable<string[]> ReadCsvRowsAsync(Stream stream, bool hasHeaderRow = false, int? rowLimit = null)
        {
            CsvHeaders = null;
            var encoding = GetEncoding(_workbook.FileEncoding);
            using var reader = new StreamReader(stream, encoding);

            int rowCount = 0;
            bool isFirstRow = true;

            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync();
                if (string.IsNullOrEmpty(line)) continue;

                var values = ParseCsvLine(line);

                if (isFirstRow && hasHeaderRow)
                {
                    CsvHeaders = values;
                    isFirstRow = false;
                    continue;
                }

                isFirstRow = false;

                if (rowLimit.HasValue && rowCount >= rowLimit.Value)
                    yield break;

                rowCount++;
                yield return values;
            }
        }

        private static string[] ParseCsvLine(string line)
        {
            var normalized = line.Replace("\", \"", "\",\"").Replace("\" ,\"", "\",\"");
            var cols = normalized.Split("\",\"");

            if (cols[0].StartsWith("\""))
                cols[0] = cols[0][1..];

            var last = cols.Length - 1;
            if (cols[last].EndsWith("\""))
                cols[last] = cols[last][..^1];

            for (int i = 0; i < cols.Length; i++)
                cols[i] = cols[i].Replace("\"\"", "\"");

            return cols;
        }

        private static Encoding GetEncoding(FileEncoding fileEncoding) => fileEncoding switch
        {
            FileEncoding.ASCII => Encoding.ASCII,
            FileEncoding.Unicode => Encoding.Unicode,
            FileEncoding.UTF32 => Encoding.UTF32,
            _ => Encoding.UTF8
        };

        public async ValueTask DisposeAsync()
        {
            // Todo: Add implementation
            await DisposeAsyncCore();

            Dispose(disposing: false);
#pragma warning disable CA1816 // Dispose methods should call SuppressFinalize
            GC.SuppressFinalize(this);
#pragma warning restore CA1816 // Dispose methods should call SuppressFinalize
        }

        protected virtual async ValueTask DisposeAsyncCore()
        {
            await Task.Delay(1);
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // dispose of anything necessary here
            }
        }
    }
}