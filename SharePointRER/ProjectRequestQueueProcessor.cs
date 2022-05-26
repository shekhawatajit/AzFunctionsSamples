using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using System.Text.Json;

namespace Onrocks.SharePoint
{
    public class ProjectRequestQueueProcessor
    {
        private readonly IPnPContextFactory pnpContextFactory;
        public ProjectRequestQueueProcessor(IPnPContextFactory pnpContextFactory)
        {
            this.pnpContextFactory = pnpContextFactory;
        }
        [FunctionName("ProjectRequestQueueProcessor")]
        public void Run([QueueTrigger("%QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo itemInfo = JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"C# Queue trigger function processed: {projectQueueItem}");
            using (var pnpContext = pnpContextFactory.Create(new Uri(itemInfo.WebUrl!)))
            {
                var requestDetails = pnpContext.Web.Lists.GetById(itemInfo.ListId).Items.GetById(itemInfo.ListItemId);
                log.LogInformation($"Title: {0}", requestDetails.FieldValuesAsText["Title"]);
                log.LogInformation($"Owners: {0}", requestDetails.FieldValuesAsText["Owners"]);
                log.LogInformation($"Members: {0}", requestDetails.FieldValuesAsText["Members"]);
                log.LogInformation($"Visitors: {0}", requestDetails.FieldValuesAsText["Visitors"]);
            }
        }
    }
}
