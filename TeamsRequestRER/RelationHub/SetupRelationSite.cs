using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Model.SharePoint;
using PnP.Framework;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using Microsoft.SharePoint.Client;


using System.Threading;
using System.Text.RegularExpressions;

namespace Adidas.OIP
{
    public class SetupRelationSite
    {
        private readonly GraphServiceClient graphClient;
        private readonly AzureFunctionSettings azureFunctionSettings;
        public SetupRelationSite(AzureFunctionSettings settings, GraphServiceClient graphServiceClient)
        {
            this.azureFunctionSettings = settings;
            this.graphClient = graphServiceClient;
        }
        [FunctionName("SetupProjectSite")]
        [FixedDelayRetry(5, "00:00:10")]
        public void Run([QueueTrigger("%RelationQueue1%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Setup of Sharepoint site process is started with data : {projectQueueItem}");
            string ProjectTitle, ProjectDescription, ProjectRequestor, teamsSiteUrl;
            using (var contextPrimaryHub = new AuthenticationManager().GetACSAppOnlyContext(info.RequestSPSiteUrl, azureFunctionSettings.ClientId, azureFunctionSettings.ClientSecret))
            {
                // Get the primary hub site details
                var primarySite = contextPrimaryHub.Site;
                contextPrimaryHub.Load(primarySite, p => p.HubSiteId, p => p.IsHubSite);

                //Reading Visitors data
                var list = contextPrimaryHub.Web.Lists.GetById(info.RequestListId);
                var requestDetails = list.GetItemById(info.RequestListItemId);
                ProjectTitle = requestDetails["Title"] == null ? string.Empty : requestDetails["Title"].ToString()!;
                ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
                ProjectRequestor = contextPrimaryHub.Web.GetUserById(info.RequestorId).UserPrincipalName;

                var owners = requestDetails.FieldValues["Owners"] as FieldUserValue[];
                var members = requestDetails.FieldValues["Members"] as FieldUserValue[];
                var visitors = requestDetails.FieldValues["Visitors"] as FieldUserValue[];

                //Generating Unique site Url
                var uniqeId = Guid.NewGuid().ToString().Split('-')[1];
                string siteName = Regex.Replace(ProjectTitle, @"\s", "") + uniqeId;

                //Reading Provisining Template Url value
                string templateUrl = Environment.GetEnvironmentVariable("ProvisioningTemplateXmlFileUrl");
                // Download the template file as stream
                var templateDocument = contextPrimaryHub.Web.GetFileByServerRelativeUrl(templateUrl);
                var downloadedContentStream = templateDocument.OpenBinaryStream();
                contextPrimaryHub.Load(templateDocument);
                contextPrimaryHub.ExecuteQuery();
                var provisioningTemplate = XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(downloadedContentStream.Value);
                downloadedContentStream.Value.Close();
                // Updating template to connect with hub
                provisioningTemplate.WebSettings.HubSiteUrl = info.RequestSPSiteUrl;
                string HubSiteIdValue = string.Format("[&quot;{0}&quot;]", primarySite.HubSiteId);
                provisioningTemplate.PropertyBagEntries.Add(new PnP.Framework.Provisioning.Model.PropertyBagEntry
                {
                    Key = "RelatedHubSiteIds",
                    Overwrite = false,
                    Value = HubSiteIdValue
                });

                //Creating new request for Teams site without Group  
                  SiteCreationOptions siteCreationOptions = new SiteCreationOptions()
              

                var teamsSiteToCreate = new SiteCreationProperties { Url = $"https://{contextPrimaryHub.Uri.DnsSafeHost}/sites/{siteName}", Owner = ProjectRequestor };

                /*  var teamsSiteToCreate = new TeamSiteWithoutGroupOptions(new Uri($"https://{contextPrimaryHub.Uri.DnsSafeHost}/sites/{siteName}"), ProjectTitle)
                  {
                      Description = ProjectDescription,
                      //Language = Language.English,
                      Owner = $"i:0#.f|membership|{ProjectRequestor}"
                  };*/
                teamsSiteUrl = teamsSiteToCreate.Url.AbsoluteUri;
                log.LogInformation($"Creating site: {teamsSiteToCreate.Url}");

                // Working on Teams Site
                using (var context = new AuthenticationManager().GetACSAppOnlyContext(TeamSiteUrl, azureFunctionSettings.ClientId, azureFunctionSettings.ClientSecret))
                {
                    // Lading site details
                    var assocSite = context.Site;
                    context.Load(assocSite, p => p.HubSiteId, p => p.IsHubSite, p => p.Id);
                    context.ExecuteQuery();

                    //Adding visitors
                    var assocWeb = context.Web;
                    context.Load(assocWeb, p => p.AssociatedVisitorGroup, p => p.Title);
                    context.ExecuteQuery();
                    // Using the value when not cleared
                    if (requestDetails["Visitors"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                        {
                            var usr = context.Web.EnsureUser(user.Sip);
                            context.Web.AssociatedVisitorGroup.Users.AddUser(usr);
                        }
                    }
                    context.ExecuteQuery();

                    // Use CSOM to load the web title
                    context.RequestTimeout = Timeout.Infinite;
                    context.ExecuteQueryRetry();
                    ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                    {
                        ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                        {
                            log.LogInformation(String.Format("{0:00}/{1:00} - {2}", progress, total, message));
                        },
                        IgnoreDuplicateDataRowErrors = true
                    };
                    assocWeb.ApplyProvisioningTemplate(provisioningTemplate, ptai);
                }
            }
        }
    }
}