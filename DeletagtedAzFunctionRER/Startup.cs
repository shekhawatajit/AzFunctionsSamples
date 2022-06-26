using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Framework.RER.Common.Tokens;
using System;
using Microsoft.Azure.WebJobs;
using Azure.Security.KeyVault.Secrets;
using Azure.Identity;
using PnP.Framework.RER.Common.Model;
[assembly: FunctionsStartup(typeof(PnP.Framework.RER.Functions.Startup))]
namespace PnP.Framework.RER.Functions
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
             builder.Services.AddSingleton(option =>
            {
                var sharepointCreds = new SharePointAppCreds();

                sharepointCreds.ClientId = azureFunctionSettings.SPClientId;
                sharepointCreds.ClientSecret = LoadSecret(azureFunctionSettings.KeyVaultName, azureFunctionSettings.SPSecretName).Value;
                return sharepointCreds;
            });
 
            builder.Services.AddSingleton<TokenManagerFactory>();
        }
        private static KeyVaultSecret LoadSecret(string KeyVaultName, string SecretName)
        {
            var KeyVaultUrl = string.Format("https://{0}.vault.azure.net/", KeyVaultName);
            SecretClient client = new SecretClient(new Uri(KeyVaultUrl), new DefaultAzureCredential());
            return client.GetSecret(SecretName).Value;
        }
    }
}
