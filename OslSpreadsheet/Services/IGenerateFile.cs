using OoxSpreadsheet.Models.Files.xlsx;

namespace OoxSpreadsheet.Services
{
    internal interface IGenerateFile : IDisposable, IAsyncDisposable
    {
        Task<byte[]> GenerateFileAsync(ooxContentTypes types, ooxWorkbook workbook);
    }
}
