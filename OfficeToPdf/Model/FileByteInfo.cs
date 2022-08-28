using System;
namespace OfficeToPdf
{
    public class FileByteInfo
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }

        public Byte[] Bytes { get; set; }
    }
}