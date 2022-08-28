using Microsoft.Extensions.Configuration;
namespace OfficeToPdf
{
    public class EnvironmentHelper
    {
        public static string GetEnvValue(string key)
        {
            //群集读取env.json
            var builder = new ConfigurationBuilder()
                .AddJsonFile("env.json");
            var configration = builder.Build();

            var value = configration[key];
            return value ?? string.Empty;
        }
    }
}