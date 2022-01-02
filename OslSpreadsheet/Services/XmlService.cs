using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace OoxSpreadsheet.Services
{
    internal static class XmlService
    {
        internal static async Task<byte[]> ConvertToXmlAsync(object obj)
        {
            try
            {
                // Generate serializer
                XmlSerializer serializer = new(obj.GetType());

                // Settings
                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    OmitXmlDeclaration = false
                };

                var ns = new XmlSerializerNamespaces();

                // Generate memory stream
                await using MemoryStream memoryStream = new();
                using var streamWriter = XmlTextWriter.Create(memoryStream, settings);

                serializer.Serialize(streamWriter, obj, ns);

                var file = memoryStream.ToArray();

                var result = Encoding.UTF8.GetString(file);

                Console.WriteLine(result);

                return file;
            }
            catch
            {
                return new byte[0];
            }
        }
    }
}
