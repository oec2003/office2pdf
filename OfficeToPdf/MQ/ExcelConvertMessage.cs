using System;

namespace OfficeToPdf
{
 
        [RabbitMq("ExcelConvertMessageQueue", ExchangeName = "ExcelConvertMessageExchange", IsProperties = false)]
        [Serializable]
        public class ExcelConvertMessage : LibreOfficeConvertMessage
        {
            private FileInfo _fileInfo;

            public ExcelConvertMessage(FileInfo fileInfo)
            {
                this._fileInfo = fileInfo;
            }

            public override string ToString() => base.ToString();

            public override FileInfo FileInfo => this._fileInfo;

            public override string MessageCatalogID => this._fileInfo.FileId;
        }
}