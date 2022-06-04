using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

using System.Security.Cryptography.X509Certificates;
using System;
using Microsoft.Azure.WebJobs;
using Azure.Security.KeyVault.Certificates;

using Azure.Identity;

[assembly: FunctionsStartup(typeof(Onrocks.SharePoint.Startup))]
namespace Onrocks.SharePoint
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
           

            builder.Services.AddSingleton(option =>
            {
                var CredentialOptions = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };
                // Certificate based Credential  singleton
                var clientCertCredential = new ClientCertificateCredential(azureFunctionSettings!.TenantId, azureFunctionSettings!.ClientId, LoadCertificate(azureFunctionSettings), CredentialOptions);
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                return new GraphServiceClient(clientCertCredential, scopes);
            });


        }
        private static X509Certificate2 LoadCertificate(AzureFunctionSettings settings)
        {
            var KeyVaultUrl = string.Format("https://{0}.vault.azure.net/", settings.KeyVaultName);
            CertificateClient client = new CertificateClient(new Uri(KeyVaultUrl), new DefaultAzureCredential());
            return client.DownloadCertificate(settings.CertificateName);
        }
    }
}