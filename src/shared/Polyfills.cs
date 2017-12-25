using System;
using System.IO;

namespace OpenMcdf
{
    public static class StreamExtension
    {
        public static void Close(this Stream stream)
        {
        }
    }

    public static class BinaryReaderExtension
    {
        public static void Close(this BinaryReader stream)
        {
        }
    }

    public class SerializableAttribute : Attribute
    {
    }
}
