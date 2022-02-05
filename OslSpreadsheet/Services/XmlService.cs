using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace OslSpreadsheet.Services
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
                    Indent = false,
                    OmitXmlDeclaration = false,
                    Encoding = Encoding.UTF8,
                    DoNotEscapeUriAttributes = true
                };

                var ns = new XmlSerializerNamespaces();
                ns.Add(string.Empty, string.Empty);

                // Generate memory stream
                await using MemoryStream memoryStream = new();
                using var streamWriter = XmlWriter.Create(memoryStream, settings);
                streamWriter.WriteStartDocument(true);

                serializer.Serialize(streamWriter, obj, ns);

                var file = memoryStream.ToArray();

                var result = Encoding.UTF8.GetString(file);

                Console.WriteLine(result.Replace("utf-8", "UTF-8"));
                Console.WriteLine(" ");

                return Encoding.UTF8.GetBytes(result.Replace("utf-8", "UTF-8").Replace("\" />", "\"/>"));

                //return file;
            }
            catch (Exception ex)
            {
                return new byte[0];
            }
        }

        internal static async Task<object> ConvertToObject<T>(string xml)
        {
            T retval = default;

            using (var reader = new StringReader(xml))
            {
                retval = (T)new XmlSerializer(typeof(T)).Deserialize(reader);
            }

            return retval;
        }
    }
}
