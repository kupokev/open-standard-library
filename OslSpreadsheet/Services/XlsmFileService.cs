using OslSpreadsheet.Models;

namespace OoxSpreadsheet.Services
{
    internal class XlsmFileService : IFileService
    {
        private bool disposedValue;

        public ValueTask DisposeAsync()
        {
            throw new NotImplementedException();
        }

        public async Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            throw new NotImplementedException();
        }

        public async Task<oWorkbook> GenerateModel(byte[] file)
        {
            throw new NotImplementedException();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~GenerateXlsmFileService()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
