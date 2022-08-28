using System;

namespace OfficeToPdf
{
    [RabbitMq("PptConvertMessageQueue", ExchangeName = "PptConvertMessageExchange", IsProperties = false)]
    [Serializable]
    public class PowerPointConvertMessage : LibreOfficeConvertMessage
    {
        private FileInfo _fileInfo;

        public PowerPointConvertMessage(FileInfo fileInfo)
        {
            this._fileInfo = fileInfo;
        }

        public override string ToString() => base.ToString();


        public override FileInfo FileInfo => this._fileInfo;

        public override string MessageCatalogID => this._fileInfo.FileId;


        public new bool IsTempStorage { get; set; } = false;
    }
}