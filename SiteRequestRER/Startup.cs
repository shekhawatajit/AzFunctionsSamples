using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth;
using System.Security.Cryptography.X509Certificates;
using System;
using Microsoft.Azure.WebJobs;
using Azure.Security.KeyVault.Certificates;
using Azure.Identity;
using PnP.Framework;

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
            var cert = LoadCertificate(azureFunctionSettings);
            builder.Services.AddPnPCore(options =>
            {
                // Configure an authentication provider with certificate (Required for app only)
                var authProvider = new X509CertificateAuthenticationProvider(azureFunctionSettings!.ClientId, azureFunctionSettings.TenantId, cert);
                // And set it as default
                options.DefaultAuthenticationProvider = authProvider;
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