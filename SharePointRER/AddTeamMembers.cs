using System;
using Microsoft.Azure.WebJobs;
using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;

using System.Collections.Generic;


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
        public void Run([QueueTrigger("%Step2QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Team members adding process started with data: {projectQueueItem}");
            try
            {
                using (var pnpContext = pnpContextFactory.Create(new Uri(info.WebUrl)))
                {
                    var Members = new List<ConversationMember>();
                    IList list = pnpContext.Web.Lists.GetById(info.ListId);
                    IListItem requestDetails = list.Items.GetById(info.ListItemId,
                            li => li.Title,
                            li => li.All);
                    // Using the value when not cleared
                    if (requestDetails["Owners"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Owners"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            var upnUser = pnpContext.Web.GetUserById(info.RequestorId).UserPrincipalName;
                            var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", upnUser);

                            var TeamUser = new AadUserConversationMember
                            {
                                Roles = new List<String>() { "owner" },
                                AdditionalData = new Dictionary<string, object>()
                               {{"user@odata.bind", objUser}}
                            };
                            Members.Add(TeamUser);
                        }
                    }
                    if (requestDetails["Members"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Members"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            var upnUser = pnpContext.Web.GetUserById(info.RequestorId).UserPrincipalName;
                            var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", upnUser);

                            var TeamUser = new AadUserConversationMember
                            {
                                Roles = new List<String>() { },
                                AdditionalData = new Dictionary<string, object>()
                               {{"user@odata.bind", objUser}}
                            };
                            Members.Add(TeamUser);
                        }
                    }

                    graphClient.Teams[info.TeamsId].Members.Add(Members).Request().PostAsync();

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