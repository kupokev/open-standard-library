using OoxSpreadsheet.Models.Files.xlsx;
using OoxSpreadsheet.Services;
using OslSpreadsheet.Models;

namespace OoxSpreadsheet
{
    public interface ISpreadsheet : IDisposable, IAsyncDisposable
    {
        OslSpreadsheet.Models.oWorkbook Workbook { get; }

        Task<byte[]> GenerateFileAsync();
    }

    public class Spreadsheet : ISpreadsheet
    {
        private readonly OslSpreadsheet.Models.oWorkbook _workbook;

        public Spreadsheet()
        {
            _workbook = new OslSpreadsheet.Models.oWorkbook();
        }
        
        public OslSpreadsheet.Models.oWorkbook Workbook  { get => _workbook; }


        public async Task<byte[]> GenerateFileAsync()
        {
            return await GenerateFileAsync(SpreadsheetFileFormat.XLSX);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public async Task<byte[]> GenerateFileAsync(SpreadsheetFileFormat fileFormat)
        {
            IGenerateFile _fileService;

            switch (fileFormat)
            {
                case SpreadsheetFileFormat.ODS:
                    _fileService = new GenerateXlsxFileService();

                    throw new NotImplementedException();

                case SpreadsheetFileFormat.XLSM:
                    _fileService = new GenerateXlsxFileService();

                    throw new NotImplementedException();

                default:
                case SpreadsheetFileFormat.XLSX:
                    _fileService = new GenerateXlsxFileService();
                    break;
            }

            //return await _fileService.GenerateFileAsync(_workbook);
            return new byte[0];
        }

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