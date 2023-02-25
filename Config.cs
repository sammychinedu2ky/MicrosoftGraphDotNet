using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace MicrosoftGraphDotNet
{
    internal class Config
    {
        public string? AzureClientId { get; set; }
        public string? AzureClientSecret { get; set; }
        public string? AzureTenantId { get; set; }
        public string? GitHubClientSecret { get; set; }

        public string? NameOfNewFile { get; set; }
        public Config()
        {
            IConfiguration config = new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json").Build();
            config.Bind(this);

        }
    }
}
