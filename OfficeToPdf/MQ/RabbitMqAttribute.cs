using System;

namespace OfficeToPdf
{
    public class RabbitMqAttribute : Attribute
    {
        public RabbitMqAttribute(string queueName) => this.QueueName = queueName ?? string.Empty;

        public string ExchangeName { get; set; }

        public string QueueName { get; private set; }

        public bool IsProperties { get; set; }

        public string DeadLetterExchange { get; set; }

        public string DeadLetterRoutingKey { get; set; }
    }
}