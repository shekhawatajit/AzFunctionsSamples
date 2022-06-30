using System;
using System.IO;
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
using

namespace SimplifiedDelegatedRER
{
    public class ListData
    {
        public string Title { get; set; }
    }
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
            ProjectRequestInfo info = new ProjectRequestInfo();
            dynamic data = JsonConvert.DeserializeObject(request.Content.ReadAsStringAsync().Result);
            log.LogInformation(request.Content.ReadAsStringAsync().Result);
            info.RequestListItemId = data?.RequestListItemId;
            info.RequestorId = data?.RequestorId;
            //Creating PnP.Core context using clientid and client secret with user imperssionation
            var secretKV = LoadSecret(_functionSettings.KeyVaultName, _functionSettings.SecretName);
            var clientSecret = new SecureString();
            foreach (char c in secretKV) clientSecret.AppendChar(c);
            var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_functionSettings.ClientId, _functionSettings.TenantId, clientSecret, () => request.Headers.Authorization.Parameter);
            using (PnPContext pnpCoreContext = await _pnpContextFactory.CreateAsync(new System.Uri(_functionSettings.HubSite), onBehalfAuthProvider))
            {
                //Creating Graph Client
                var graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
                {
                    return pnpCoreContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), requestMessage);
                }));

                //Creating Teams
                TeamsHelper tm = new TeamsHelper(pnpCoreContext, graphServiceClient, log);
                info.TeamsId = tm.CreateTeams(info);

            }

            /* 
             var query = request.RequestUri.ParseQueryString();
             var siteUrl = query["siteUrl"];
             var tenantId = query["tenantId"];
 */
            /*var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(requestBody);

                string json = JsonConvert.SerializeXmlNode(xmlDoc);
                JObject eventData = JObject.Parse(json);

                this._info.RequestListItemId = (int)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListItemId"];
                this._info.RequestListId = Guid.Parse((string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListId"]);
                this._info.RequestSPSiteUrl = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["WebUrl"];
                this._info.RequestorId = (int)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["CurrentUserId"];
                // Important: Here we are laoding SharePOint App Client Secret
                this._info.SPSecret = LoadSecret(_functionSettings.KeyVaultName, _functionSettings.SPSecretName);
                this._info.ContextToken = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ContextToken"];

                var host = req.Host.Host;
                var tokenManager = new TokenManager(_functionSettings.SPClientId, _info.SPSecret, _client, _info.ContextToken, host);

                var context = await tokenManager.GetUserClientContextAsync(_info.RequestSPSiteUrl);
                var userToekn = await tokenManager.GetAccessTokenAsync(_info.RequestSPSiteUrl);
                context.Load(context.Web);
                await context.ExecuteQueryAsync();



                //Creating PnP Core Context using Delegated Permission
                //var clientSecret = new SecureString();
                // Important: Here we are laoding Azure App Client Secret
                // var AppSecret = LoadSecret(_functionSettings.KeyVaultName, _functionSettings.SecretName);
                //foreach (char c in AppSecret) clientSecret.AppendChar(c);
                // var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_functionSettings.ClientId, _functionSettings.TenantId, clientSecret, () => userToekn);

                // Creating Graph Service Client using Delegated permission


                using (PnPContext pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(context))
                {
                    await pnpCoreContext.Web.LoadAsync(w => w.Title, w => w.Id); // HTTP request is executed immediately
                    log.LogInformation("pnpCoreContext.Web.Title");

                    log.LogInformation(pnpCoreContext.Web.Title);
                    // Teams tm = new Teams(pnpContext, )

                    //Creating Graph Client
                    var graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
                    {
                        return pnpCoreContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com"), requestMessage);
                    }));
                }
            }

            catch (System.Exception err)
            {
                log.LogError(err.ToString());
                responseMessage = err.Message;
            }*/
            return new JsonResult(info);
        }
        private string LoadSecret(string KeyVaultName, string SecretName)
        {
            var KeyVaultUrl = string.Format("https://{0}.vault.azure.net/", KeyVaultName);
            SecretClient client = new SecretClient(new Uri(KeyVaultUrl), new DefaultAzureCredential());
            return client.GetSecret(SecretName).Value.Value;
        }
    }
}
