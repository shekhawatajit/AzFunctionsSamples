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

namespace SimplifiedDelegatedRER
{
    public class ProjectRequestAdded
    {
        private readonly AzureFunctionSettings _functionSettings;
        private ProjectRequestInfo _info = new ProjectRequestInfo();
        private readonly IPnPContextFactory _pnpContextFactory;
        public ProjectRequestAdded(AzureFunctionSettings azureFunctionSettings, IPnPContextFactory pnpContextFactory)
        {
            _functionSettings = azureFunctionSettings;
            this._pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("ProjectRequestAdded")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "ProjectRequestAdded")] HttpRequestMessage request, ILogger log)
        {
            log.LogInformation("Item Added HTTP trigger function processed a request.");

            //Processing request body
            ProjectRequestInfo info = JsonSerializer.Deserialize<ProjectRequestInfo>(request.Content.ReadAsStringAsync().Result);
            Utilities ut = new Utilities();
            //Creating PnP.Core context using clientid and client secret with user imperssionation
            var secretKV = ut.LoadSecret(_functionSettings.KeyVaultName, _functionSettings.SecretName);
            var clientSecret = new SecureString();
            foreach (char c in secretKV) clientSecret.AppendChar(c);
            var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_functionSettings.ClientId, _functionSettings.TenantId, clientSecret, () => request.Headers.Authorization.Parameter);
            using (PnPContext pnpCoreContext = await _pnpContextFactory.CreateAsync(new System.Uri(_functionSettings.OIPHubSite), onBehalfAuthProvider))
            {
                //Creating Graph Client
                var graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
                {
                    return pnpCoreContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), requestMessage);
                }));

                //Creating Teams
                TeamsHelper tm = new TeamsHelper(pnpCoreContext, graphServiceClient, log);
                info.TeamsId = tm.CreateTeams(info);
                TeamSiteHelper tspHelper = new TeamSiteHelper(pnpCoreContext, graphServiceClient, log, _functionSettings);
            }
            return new OkObjectResult("OK");
        }
    }
}
