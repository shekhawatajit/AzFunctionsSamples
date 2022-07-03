using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using PnP.Core.Services;
using System.Security;
using PnP.Core.Auth;
using System.Text.Json;
using PnP.Framework.Provisioning.Model;

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

            var jsonString = System.Text.Json.JsonSerializer.Serialize(info);
            log.LogInformation(jsonString);

            //Creating PnP.Core context using clientid and client secret with user imperssionation
            var secretKV = ut.LoadSecret(_functionSettings.KeyVaultName, _functionSettings.SecretName);
            var clientSecret = new SecureString();
            foreach (char c in secretKV) clientSecret.AppendChar(c);
            var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_functionSettings.ClientId, _functionSettings.TenantId, clientSecret, () => request.Headers.Authorization.Parameter);
            ProvisioningTemplate template;
            //Working on hub site for reading required informaiton
            using (PnPContext contextPrimaryHub = await _pnpContextFactory.CreateAsync(new System.Uri(info.RequestSPSiteUrl), onBehalfAuthProvider))
            {

                //Reading Project request details from SharePoint
                info = await ut.ReadRequestFromList(contextPrimaryHub, info, log);

                //Reading ProvisionTemplate
                template = await ut.ReadProvisionTemplte(contextPrimaryHub, info.ProvisionTemplate, log);

                ut.UpdateSpList(contextPrimaryHub, info, _functionSettings, log);
            }

            // Working on New Teams Site
            using (PnPContext newTeamsSiteContext = await _pnpContextFactory.CreateAsync(new System.Uri(info.NewSiteUrl), onBehalfAuthProvider))
            {
                switch (info.SiteType)
                {
                    case "GroupWithTeams":
                        // Applying provising template
                        //System.Threading.Thread.Sleep(5000);
                        // Provision first so that durnig this time Group and Teams will be ready
                        ut.ProvisionSite(newTeamsSiteContext, info, template, log);
                        // Creating Teams from SharePoint Team Site
                        //System.Threading.Thread.Sleep(5000);
                        await ut.CreateTeamsFromSPSite(newTeamsSiteContext, log);
                        //Adding membersand owners to team
                        System.Threading.Thread.Sleep(5000);
                        await ut.AddTeamMembers(newTeamsSiteContext, info, log);
                        break;
                    case "GroupWithoutTeams":
                        // Applying provising template
                        //System.Threading.Thread.Sleep(5000);
                         // Provision first so that durnig this time Group and Teams will be ready
                        ut.ProvisionSite(newTeamsSiteContext, info, template, log);
                        //Adding members
                        //System.Threading.Thread.Sleep(5000);
                        await ut.AddSiteMembers(newTeamsSiteContext, info, log);
                        break;
                    default:
                        log.LogError("Invalid Site Type. Allowed 'GroupWithTeams' or 'GroupWithoutTeams'");
                        break;
                }
            }
            return new OkObjectResult("OK");
        }
    }
}
