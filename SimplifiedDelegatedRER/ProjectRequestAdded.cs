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
                template = await ut.ReadProvisionTemplte(contextPrimaryHub, log, info.ProvisionTemplate);

                ut.UpdateSpList(_functionSettings.MailListTitle, info.ProjectTitle, info.ProjectDescription, info.ProjectRequestor, info.NewSiteUrl, contextPrimaryHub);
            }

            // Working on New Teams Site
            using (PnPContext newTeamsSiteContext = await _pnpContextFactory.CreateAsync(new System.Uri(info.NewSiteUrl), onBehalfAuthProvider))
            {
                switch (info.SiteType)
                {
                    case "GroupWithTeams":
                        // Creating Teams from SharePoint Team Site
                        await ut.CreateTeamsFromSPSite(newTeamsSiteContext, log);
                        // Applying provising template
                        ut.ProvisionSite(newTeamsSiteContext, log, template, info);
                        //Adding membersand owners to team
                        await ut.AddTeamMembers(newTeamsSiteContext, info, log);
                        break;
                    case "GroupWithoutTeams":
                        // Applying provising template
                        ut.ProvisionSite(newTeamsSiteContext, log, template, info);
                        //Adding members
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
