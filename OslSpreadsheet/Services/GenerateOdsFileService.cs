using OslSpreadsheet.Models;
using OslSpreadsheet.Models.Files.ods;
using OslSpreadsheet.Services;
using System.Text;

namespace OoxSpreadsheet.Services
{
    internal class GenerateOdsFileService : IGenerateFile
    {
        private bool disposedValue;

        public async ValueTask DisposeAsync()
        {
            await Task.Run(() => Dispose());
        }

        public async Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            byte[] output;

            try
            {             
                ODManifest manifest = new();
                ODContent content = new();
                ODMeta meta = new();
                ODStyles style = new();

                // Add models to a file list for compression
                List<InMemoryFile> files = new()
                {
                    new InMemoryFile()
                    {
                        FileName = "META-INF/manifext.xml",
                        Content = await XmlService.ConvertToXmlAsync(manifest)
                    },
                    new InMemoryFile()
                    {
                        FileName = "content.xml",
                        Content = await XmlService.ConvertToXmlAsync(content)
                    },
                    new InMemoryFile()
                    {
                        FileName = "meta.xml",
                        Content = await XmlService.ConvertToXmlAsync(meta)
                    },
                    new InMemoryFile()
                    {
                        FileName = "mimetype",
                        //Content = await XmlService.ConvertToXmlAsync(mimeType)
                        Content = Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.spreadsheet")
                    },
                    new InMemoryFile()
                    {
                        FileName = "style.xml",
                        Content = await XmlService.ConvertToXmlAsync(style)
                    }
                };

                output = await ZipService.GenerateZipAsync(files);
            }
            catch
            {
                throw new Exception("There was an issue compressing the file.");
            }

            return output;
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
        // ~GenerateOdsFileService()
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
