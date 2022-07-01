using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using Azure.Security.KeyVault.Secrets;
using Azure.Identity;
using PnP.Core.Services;
using System.Security;
using PnP.Core.Auth;
using Microsoft.Graph;
using System.Text.Json;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Framework;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using Microsoft.SharePoint.Client;
using System.Threading;
using System.IO;
using PnP.Core.QueryModel;
using System.Linq;
using System.Collections.Generic;

namespace SimplifiedDelegatedRER
{
    public class RelationRequestAdded
    {
        private readonly AzureFunctionSettings _functionSettings;
        private ProjectRequestInfo _info = new ProjectRequestInfo();
        private readonly IPnPContextFactory _pnpContextFactory;
        public RelationRequestAdded(AzureFunctionSettings azureFunctionSettings, IPnPContextFactory pnpContextFactory)
        {
            _functionSettings = azureFunctionSettings;
            this._pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("RelationRequestAdded")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "RelationRequestAdded")] HttpRequestMessage request, ILogger log)
        {
            log.LogInformation("Item Added HTTP trigger function processed a request.");

            //Processing request body
            RelationRequestInfo info = JsonSerializer.Deserialize<RelationRequestInfo>(request.Content.ReadAsStringAsync().Result);
            Utilities ut = new Utilities();

            //Creating PnP.Core context using clientid and client secret with user imperssionation
            var secretKV = ut.LoadSecret(_functionSettings.KeyVaultName, _functionSettings.SecretName);
            var clientSecret = new SecureString();
            foreach (char c in secretKV) clientSecret.AppendChar(c);
            var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_functionSettings.ClientId, _functionSettings.TenantId, clientSecret, () => request.Headers.Authorization.Parameter);
            using (PnPContext pnpCoreContext = await _pnpContextFactory.CreateAsync(new System.Uri(_functionSettings.RelationHubSite), onBehalfAuthProvider))
            {
                //Creating Graph Client
                var graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
                {
                    return pnpCoreContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), requestMessage);
                }));

                //Creating SharePOint Site
                SharePointSiteHelper tm = new SharePointSiteHelper(pnpCoreContext, graphServiceClient, log, _functionSettings);
                info.RelationSiteUrl = await tm.SetupSite(info, ut);
            }
            return new OkObjectResult("OK");
        }
    }
}
