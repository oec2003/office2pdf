using System;

namespace OfficeToPdf
{
    [RabbitMq("WordConvertMessageQueue", ExchangeName = "WordConvertMessageExchange", IsProperties = false)]
    [Serializable]
    public class WordConvertMessage : LibreOfficeConvertMessage
    {
        private FileInfo _fileInfo;

        public WordConvertMessage(FileInfo fileInfo)
        {
            this._fileInfo = fileInfo;
        }

        public override FileInfo FileInfo => this._fileInfo;

        public override string MessageCatalogID => this._fileInfo.FileId;

        public new bool IsTempStorage { get; set; } = false;
    }
}