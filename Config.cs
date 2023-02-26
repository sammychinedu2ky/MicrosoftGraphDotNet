using Microsoft.Extensions.Configuration;

namespace MicrosoftGraphDotNet
{
    internal class Config
    {
        // Define properties to hold configuration values
        public string? AzureClientId { get; set; }
        public string? AzureClientSecret { get; set; }
        public string? AzureTenantId { get; set; }
        public string? GitHubClientSecret { get; set; }
        public string? NameOfNewFile { get; set; }

        // Constructor to read configuration values from appsettings.json file
        public Config()
        {
            // Create a new configuration builder and add appsettings.json as a configuration source
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();

            // Bind configuration values to the properties of this class
            config.Bind(this);
        }
    }
}
