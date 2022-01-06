using OslSpreadsheet.Models;

namespace OoxSpreadsheet.Services
{
    internal interface IGenerateFile : IDisposable, IAsyncDisposable
    {
        Task<byte[]> GenerateFileAsync(oWorkbook workbook);
    }
}
