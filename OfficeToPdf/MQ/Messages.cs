namespace OfficeToPdf
{
    public abstract class Messages
    {
        public string Queues { get; set; }

        public string Exchange { get; set; }
        
        public abstract string MessageCatalogID { get; }
    }
}