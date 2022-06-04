using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Threading.Tasks;
using PnP.Core.Model.Security;
using System.Collections.Generic;
namespace Onrocks.SharePoint
{
    public class SetupProjectSite
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly GraphServiceClient graphClient;
        public SetupProjectSite(IPnPContextFactory pnpContextFactory, GraphServiceClient graphServiceClient)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.graphClient = graphServiceClient;
        }
        [FunctionName("SetupProjectSite")]
        [FixedDelayRetry(5, "00:00:10")]
        public async Task Run([QueueTrigger("%Step3QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Setup of Sharepoint Teams site process is started with data : {projectQueueItem}");

            //Get Teams site Url 
            var TeamSiteUrl = graphClient.Groups[info.TeamsId].Sites["root"].Request().GetAsync().Result.WebUrl;

            using (var contextPrimaryHub = await pnpContextFactory.CreateAsync(new Uri(info.RequestSPSiteUrl)))
            {
                // Get the primary hub site details
                ISite primarySite = await contextPrimaryHub.Site.GetAsync(
                    p => p.HubSiteId,
                    p => p.IsHubSite);

                //Reading Visitors data
                IList list = contextPrimaryHub.Web.Lists.GetById(info.RequestListId);
                IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                        li => li.Title,
                        li => li.All);

                // Working on Teams Site
                using (var context = await pnpContextFactory.CreateAsync(new Uri(TeamSiteUrl)))
                {
                    // Associating to Hub
                    ISite assocSite = await context.Site.GetAsync(
                        p => p.HubSiteId,
                        p => p.IsHubSite);
                    if (assocSite.HubSiteId == Guid.Empty)
                    {
                        var resultJoin = await assocSite.JoinHubSiteAsync(primarySite.HubSiteId);
                        log.LogInformation($"Site connected to Hub: {resultJoin}");
                    }

                    //Adding visitors
                    await context.Web.LoadAsync(p => p.AssociatedVisitorGroup);
                    // Using the value when not cleared
                    if (requestDetails["Visitors"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                            if (LoginName != null)
                            {
                                context.Web.AssociatedVisitorGroup.AddUser(LoginName);
                            }
                        }
                    }
                }
            }
        }
    }
}