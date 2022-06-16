using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Framework;
using System;
using Microsoft.Azure.WebJobs;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Graph;
using Azure.Identity;

[assembly: FunctionsStartup(typeof(Adidas.OIP.Startup))]
namespace Adidas.OIP
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var azureFunctionSettings = new AzureFunctionSettings();
            builder.Services.AddSingleton(option =>
            {
                var config = builder.GetContext().Configuration;
                config.Bind(azureFunctionSettings);
                azureFunctionSettings.ClientSecret = LoadSecret(azureFunctionSettings).Value;
                return azureFunctionSettings;
            });
            builder.Services.AddSingleton(option =>
            {
                var CredentialOptions = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };
                // Certificate based Credential  singleton
                var clientSecretCredential = new ClientSecretCredential(azureFunctionSettings!.TenantId, azureFunctionSettings!.ClientId, azureFunctionSettings!.ClientSecret, CredentialOptions);
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                return new GraphServiceClient(clientSecretCredential, scopes);
            });
        }
        private static KeyVaultSecret LoadSecret(AzureFunctionSettings settings)
        {
            var KeyVaultUrl = string.Format("https://{0}.vault.azure.net/", settings.KeyVaultName);
            SecretClient client = new SecretClient(new Uri(KeyVaultUrl), new DefaultAzureCredential());
            return client.GetSecret(settings.SecretName).Value;
        }
    }
}