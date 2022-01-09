using OslSpreadsheet.Models;

namespace OoxSpreadsheet.Services
{
    internal interface IFileService : IDisposable, IAsyncDisposable
    {
        Task<byte[]> GenerateFileAsync(oWorkbook workbook);

        Task<oWorkbook> GenerateModel(byte[] file);
    }
}
