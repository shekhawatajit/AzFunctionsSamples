using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using PnP.Core.Admin.Model.SharePoint;

namespace Onrocks.SharePoint
{
    public class SetupSite
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly GraphServiceClient graphClient;
        public SetupSite(IPnPContextFactory pnpContextFactory, GraphServiceClient graphServiceClient)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.graphClient = graphServiceClient;
        }
        [FunctionName("SetupSite")]
        [FixedDelayRetry(5, "00:00:10")]
        public async Task Run([QueueTrigger("%Step1QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Setup of Sharepoint site process is started with data : {projectQueueItem}");
            try
            {
                string ProjectTitle, ProjectDescription, ProjectRequestor;
                using (var contextPrimaryHub = await pnpContextFactory.CreateAsync(new Uri(info.RequestSPSiteUrl)))
                {
                    //Reading data from SharePoint list
                    // Get the primary hub site details
                    ISite primarySite = await contextPrimaryHub.Site.GetAsync(
                        p => p.HubSiteId,
                        p => p.IsHubSite);

                    IList list = contextPrimaryHub.Web.Lists.GetById(info.RequestListId);
                    IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                            li => li.Title,
                            li => li.All);
                    ProjectTitle = requestDetails.Title == null ? string.Empty : requestDetails.Title;
                    ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
                    ProjectRequestor = contextPrimaryHub.Web.GetUserById(info.RequestorId).UserPrincipalName;
                    string siteName = Regex.Replace(ProjectTitle, @"\s", "");

                    //Creating new request for Teams site without Group  
                    var teamsSiteToCreate = new TeamSiteWithoutGroupOptions(new Uri($"https://{contextPrimaryHub.Uri.DnsSafeHost}/sites/{siteName}"), ProjectTitle)
                    {
                        Description = ProjectDescription,
                        Language = Language.English,
                        Owner = $"i:0#.f|membership|{ProjectRequestor}"
                    };
                    log.LogInformation($"Creating site: {teamsSiteToCreate.Url}");

                    // Create the new site collection
                    using (var newSiteContext = await contextPrimaryHub.GetSiteCollectionManager().CreateSiteCollectionAsync(teamsSiteToCreate))
                    {
                        log.LogInformation($"Site created: {teamsSiteToCreate.Url}");
                        // Associating to Hub
                        ISite assocSite = await newSiteContext.Site.GetAsync(
                            p => p.HubSiteId,
                            p => p.IsHubSite);
                        if (assocSite.HubSiteId == Guid.Empty)
                        {
                            var resultJoin = await assocSite.JoinHubSiteAsync(primarySite.HubSiteId);
                            log.LogInformation($"Site connected to Hub: {resultJoin}");
                        }

                        //Adding Users
                        await newSiteContext.Web.LoadAsync(p => p.AssociatedOwnerGroup, p => p.AssociatedMemberGroup, p => p.AssociatedVisitorGroup);
                        // Using the value when not cleared
                        if (requestDetails["Owners"] != null)
                        {
                            foreach (IFieldUserValue user in (requestDetails["Owners"] as IFieldValueCollection)!.Values)
                            {
                                // Get the stored user lookup id value
                                var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                                if (LoginName != null)
                                {
                                    newSiteContext.Web.AssociatedOwnerGroup.AddUser(LoginName);
                                }
                            }
                        }
                        if (requestDetails["Members"] != null)
                        {
                            foreach (IFieldUserValue user in (requestDetails["Members"] as IFieldValueCollection)!.Values)
                            {
                                // Get the stored user lookup id value
                                var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                                if (LoginName != null)
                                {
                                    newSiteContext.Web.AssociatedMemberGroup.AddUser(LoginName);
                                }
                            }
                        }
                        if (requestDetails["Visitors"] != null)
                        {
                            foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                            {
                                // Get the stored user lookup id value
                                var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                                if (LoginName != null)
                                {
                                    newSiteContext.Web.AssociatedVisitorGroup.AddUser(LoginName);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
    }
}