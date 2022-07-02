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
        private Utilities ut = new Utilities();
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
            info.ProjectTitle = string.Format("Protect {0}", info.RequestListItemId);
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
                    return pnpCoreContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com/v1.0"), requestMessage);
                }));

                //Creating Teams
                TeamsHelper tm = new TeamsHelper(pnpCoreContext, graphServiceClient, log);
                info = tm.CreateTeams(info);
                TeamSiteHelper tspHelper = new TeamSiteHelper(pnpCoreContext, graphServiceClient, log, _functionSettings);


                int maxRetry = 60;
                string TeamSiteUrl = string.Empty;
                //This retry to get SharePoint site url is VERY BAD but no other option left
                for (int tryNumber = 1; tryNumber <= maxRetry; tryNumber++)
                {
                    try
                    {
                        //Get Teams site Url 
                        //requried Permission: SharePoint -> Sites.FullControl.All
                        TeamSiteUrl = graphServiceClient.Groups[info.TeamsId].Sites["root"].Request().GetAsync().Result.WebUrl;
                        log.LogInformation($"Teams site Url to process : {TeamSiteUrl}");
                        log.LogInformation($"Find Url in try number : {tryNumber}");
                        // Breaking the loop on success
                        break;

                    }
                    catch (System.Exception err)
                    {
                        System.Threading.Thread.Sleep(1000);
                        log.LogInformation(err.Message);
                    }
                }
                info.ProjectSite = TeamSiteUrl;
                log.LogInformation($"Outside retry {info.ProjectSite}");
                using (PnPContext pnpProjectContext = await _pnpContextFactory.CreateAsync(new System.Uri(info.ProjectSite), onBehalfAuthProvider))
                {
                    log.LogInformation($"Inside context {pnpProjectContext.ToString()}");
                    tspHelper.SetupTeamSite(info, pnpProjectContext);
                }
                //Sending Email to Owner
                ut.UpdateSpList(_functionSettings.OIPMailListTitle, info.ProjectTitle, info.ProjectDescription, info.ProjectRequestor, info.ProjectSite, pnpCoreContext);

            }
            return new OkObjectResult("OK");
        }
    }
}
