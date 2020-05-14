using System.IO;
using System.IO.Compression;
using System.Xml;

namespace TranslatorLibrary
{
    public static class Compression
    {

        public static void ZipXmlFileToStream(ZipArchive archive, XmlDocument file, string name)
        {
            var fileInArchive = archive.CreateEntry(name, CompressionLevel.Optimal);
            using (var entryStream = fileInArchive.Open())
            {
                var writer = new XmlTextWriter(entryStream, System.Text.Encoding.UTF8);
                writer.Formatting = Formatting.Indented;
                file.WriteContentTo(writer);
                writer.Flush();
            }
        }
    }
}
