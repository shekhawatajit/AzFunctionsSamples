using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth;
using System.Security.Cryptography.X509Certificates;
using System;
using Microsoft.Azure.WebJobs;
using Azure.Security.KeyVault.Certificates;
using Microsoft.Graph;
using Azure.Identity;

[assembly: FunctionsStartup(typeof(Onrocks.SharePoint.Startup))]
namespace Onrocks.SharePoint
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var azureFunctionSettings = new AzureFunctionSettings();
            config.Bind(azureFunctionSettings);
            /*            #nullable enable
                        AzureFunctionSettings? azureFunctionSettings = null;
                        #nullable disable
                        // Add our global configuration instance
                        builder.Services.AddSingleton(options =>
                        {
                            var config = builder.GetContext().Configuration;
                            var azureFunctionSettings = new AzureFunctionSettings();
                            config.Bind(azureFunctionSettings);
                            return config;
                        });
                        // Add our configuration class
                        builder.Services.AddSingleton(options => { return azureFunctionSettings!; });*/
            builder.Services.AddPnPCore(options =>
            {
                // Configure an authentication provider with certificate (Required for app only)
                var authProvider = new X509CertificateAuthenticationProvider(azureFunctionSettings!.ClientId, azureFunctionSettings.TenantId, LoadCertificate(azureFunctionSettings));
                // And set it as default
                options.DefaultAuthenticationProvider = authProvider;
            });
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