using System;

namespace OfficeToPdf
{
    /// <summary>
    /// 
    /// </summary>
    public class MQConfig
    {
        public string Host { get; set; }
        public ushort HeartBeat { get; set; }
        public bool AutomaticRecoveryEnabled { get; set; }
        public TimeSpan NetworkRecoveryInterval { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Port { get; set; }

    }
}
