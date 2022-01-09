using OoxSpreadsheet.Services;
using OslSpreadsheet.Models;

namespace OoxSpreadsheet
{
    public interface ISpreadsheet : IDisposable, IAsyncDisposable
    {
        oWorkbook Workbook { get; }

        Task<byte[]> GenerateOdsFileAsync();

        Task<byte[]> GenerateXlsxFileAsync();
    }

    public class Spreadsheet : ISpreadsheet
    {
        private readonly OslSpreadsheet.Models.oWorkbook _workbook;

        public Spreadsheet()
        {
            _workbook = new OslSpreadsheet.Models.oWorkbook();
        }
        
        public OslSpreadsheet.Models.oWorkbook Workbook  { get => _workbook; }


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