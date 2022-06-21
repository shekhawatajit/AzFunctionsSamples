using System;
using Microsoft.Azure.WebJobs;
using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Text.Json;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Onrocks.SharePoint
{
    public class AddTeamMembers
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly GraphServiceClient graphClient;
        public AddTeamMembers(IPnPContextFactory pnpContextFactory, GraphServiceClient graphServiceClient)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.graphClient = graphServiceClient;
        }
        [FunctionName("AddTeamMembers")]
        [FixedDelayRetry(5, "00:00:10")]
        public async Task Run([QueueTrigger("%Step2QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Team members adding process started with data: {projectQueueItem}");
            try
            {
                using (var pnpContext = pnpContextFactory.Create(new Uri(info.RequestSPSiteUrl)))
                {
                    var Members = new List<ConversationMember>();
                    IList list = pnpContext.Web.Lists.GetById(info.RequestListId);
                    IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                            li => li.Title,
                            li => li.All);
                    // Using the value when not cleared
                    if (requestDetails["Owners"] != null)
                    {
                        //log.LogInformation("");
                        foreach (IFieldUserValue user in (requestDetails["Owners"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            //requried Permission: Microsoft Graph -> User.Read.All
                            var upnUser = pnpContext.Web.GetUserById(user.LookupId).UserPrincipalName;

                            var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", upnUser);

                            var TeamUser = new AadUserConversationMember
                            {
                                Roles = new List<String>() { "owner" },
                                AdditionalData = new Dictionary<string, object>()
                               {{"user@odata.bind", objUser}}
                            };
                            //var result = await graphClient.Teams[info.TeamsId].Members.Request().AddResponseAsync(TeamUser);
                            //log.LogInformation(result.StatusCode.ToString());
                            Members.Add(TeamUser);
                        }
                    }
                    if (requestDetails["Members"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Members"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            //requried Permission: Microsoft Graph -> User.Read.All
                            var upnUser = pnpContext.Web.GetUserById(user.LookupId).UserPrincipalName;
                            var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", upnUser);

                            var TeamUser = new AadUserConversationMember
                            {
                                Roles = new List<String>() { },
                                AdditionalData = new Dictionary<string, object>()
                               {{"user@odata.bind", objUser}}
                            };
                            //var result = await graphClient.Teams[info.TeamsId].Members.Request().AddResponseAsync(TeamUser);
                            //log.LogInformation(result.StatusCode.ToString());
                            Members.Add(TeamUser);
                        }
                    }
                    //Required Permissions:'TeamMember.ReadWrite.All'
                    var response  = await graphClient.Teams[info.TeamsId].Members.Add(Members).Request().PostAsync();
                    
                    //log.LogInformation(response.CurrentPage.  .Content.ReadAsStringAsync());

                    //Sending Teams info and request info in Queue 2
                    UpdateStep3Queue(info);
                }
            }
            catch (System.Exception err)
            {
                log.LogInformation(err.StackTrace);
            }
        }

        private void UpdateStep3Queue(ProjectRequestInfo info)
        {
            string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");

            //Sending New teams Id in next queue for step 3
            var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(info);
            string QueueName = Environment.GetEnvironmentVariable("Step3QueueName");
            QueueClient theQueue = new QueueClient(connectionString, QueueName);
            var itemInfoBytes = System.Text.Encoding.UTF8.GetBytes(jsonString);
            theQueue.SendMessage(System.Convert.ToBase64String(itemInfoBytes));
        }
    }
}