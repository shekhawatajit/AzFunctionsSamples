using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Azure.WebJobs;

[assembly: FunctionsStartup(typeof(SimplifiedDelegatedRER.Startup))]
namespace SimplifiedDelegatedRER
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            builder.Services.AddHttpClient();
            var azureFunctionSettings = new AzureFunctionSettings();
            builder.Services.AddSingleton(option =>
            {
                var config = builder.GetContext().Configuration;    
                config.Bind(azureFunctionSettings);
                return azureFunctionSettings;
            });
            builder.Services.AddPnPCore();
        }
    }
}