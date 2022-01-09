using OoxSpreadsheet.Models.Files.xlsx;
using OslSpreadsheet.Models;
using OslSpreadsheet.Services;
using System.IO.Compression;

namespace OoxSpreadsheet.Services
{
    internal class XlsxFileService : IFileService
    {
        private bool disposedValue;

        public ValueTask DisposeAsync()
        {
            throw new NotImplementedException();
        }

        public async Task<byte[]> GenerateFileAsync(oWorkbook workbook)
        {
            byte[] output;

            try
            {
                // TODO: need to convert oWorkbook to ooxWorkbook
                OXWorkbook wb = new(); 

                var types = new OXContentTypes(wb);

                List<InMemoryFile> files = new();

                var fileLists = types.Parts.Select(x => x.PartName).ToList();

                files.Add(new InMemoryFile()
                {
                    FileName = "[Content_Types].xml",
                    Content = await XmlService.ConvertToXmlAsync(types)
                });

                files.Add(new InMemoryFile()
                {
                    FileName = "_rels/.rels.xml",
                    Content = await XmlService.ConvertToXmlAsync(new Rels())
                });

                foreach (var file in fileLists)
                {
                    object? o = null;

                    if (file.Contains("/workbook.xml")) o = wb;
                    if (file.Contains("/xl/worksheets/")) o = wb.Sheets.First(x => x.name == file.Replace("/xl/worksheets/", "").Replace(".xml", ""));

                    files.Add(new InMemoryFile()
                    {
                        FileName = file.TrimStart('/'),
                        Content = (o == null) ? new byte[0] : await XmlService.ConvertToXmlAsync(o)
                    });
                }

                await using MemoryStream archiveStream = new();
                using (ZipArchive archive = new(archiveStream, ZipArchiveMode.Create, true))
                {
                    foreach (var file in files)
                    {
                        var zipArchiveEntry = archive.CreateEntry(file.FileName, CompressionLevel.Fastest);
                        using var zipStream = zipArchiveEntry.Open();
                        zipStream.Write(file.Content, 0, file.Content.Length);
                    }
                }

                output = archiveStream.ToArray();

                //output = await ZipService.GenerateZipAsync(files);
            }
            catch
            {
                throw new Exception("There was an issue compressing the file.");
            }

            return output;
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
        // ~GenerateXlsxFileService()
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
