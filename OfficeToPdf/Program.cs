using System;
using System.Threading;
using System.Configuration;
using NLog;
using RabbitMQ.Client;

namespace OfficeToPdf
{
    class Program
    {
        static IOfficeConverter _converter = new OfficeConverter();
        static Logger _logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            try
            {
                var mqManager = new MQManager(new MQConfig
                {
                    AutomaticRecoveryEnabled = true,
                    HeartBeat = 60,
                    NetworkRecoveryInterval = new TimeSpan(60),
                    Host = EnvironmentHelper.GetEnvValue("MQHostName"),
                    UserName = EnvironmentHelper.GetEnvValue("MQUserName"),
                    Password = EnvironmentHelper.GetEnvValue("MQPassword"),
                    Port = EnvironmentHelper.GetEnvValue("MQPort")
                });
                
                if (mqManager.Connected)
                {
                    _logger.Log(LogLevel.Info, "RabbitMQ连接成功。");
                    _logger.Log(LogLevel.Info, "RabbitMQ消息接收中...");

                    mqManager.Subscribe<PowerPointConvertMessage>(Convert);
                    mqManager.Subscribe<WordConvertMessage>(Convert);
                    mqManager.Subscribe<ExcelConvertMessage>(Convert);
                }
                else
                {
                    _logger.Warn("RabbitMQ连接初始化失败,请检查连接。");
                    Console.ReadLine();
                }
            }
            catch(Exception ex)
            {
                _logger.Error(ex.Message);
            }
        }

        private static readonly object Lock = new object();

        static void Convert(LibreOfficeConvertMessage message)
        {
            lock (Lock)
            {
                if (message != null)
                {
                    _converter.OnWork(message);
                    _logger.Log(LogLevel.Info, "接受文件信息：" + Newtonsoft.Json.JsonConvert.SerializeObject(message.FileInfo));
                }
            }
        }
    }
}
