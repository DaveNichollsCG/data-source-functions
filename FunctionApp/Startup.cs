using System;
using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Plumsail.DataSource.FunctionApp;
using SP = Plumsail.DataSource.SharePoint;

[assembly: FunctionsStartup(typeof(Startup))]
namespace Plumsail.DataSource.FunctionApp
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            IConfigurationRoot configuration = new ConfigurationBuilder()
              .SetBasePath(Environment.CurrentDirectory)
              .AddJsonFile("appsettings.local.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables()
              .Build();

            builder.Services.Configure<SP.Settings.AppSettings>(configuration);
            builder.Services.AddTransient<SP.GraphServiceClientProvider>();
        }
    }
}
