using System;
using Microsoft.Azure.WebJobs;
using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Framework;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;

namespace Adidas.OIP
{
    public class AddTeamMembers
    {
        private readonly GraphServiceClient graphClient;
        private readonly AzureFunctionSettings azureFunctionSettings;
        public AddTeamMembers(AzureFunctionSettings settings, GraphServiceClient graphServiceClient)
        {
            this.azureFunctionSettings = settings;
            this.graphClient = graphServiceClient;
        }
        [FunctionName("AddTeamMembers")]
        [FixedDelayRetry(5, "00:00:10")]
        public void Run([QueueTrigger("%Step2QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Team members adding process started with data: {projectQueueItem}");
            try
            {
                using (var pnpContext = new AuthenticationManager().GetACSAppOnlyContext(info.RequestSPSiteUrl, azureFunctionSettings.ClientId, azureFunctionSettings.ClientSecret))
                {
                    var Members = new List<ConversationMember>();
                    var list = pnpContext.Web.Lists.GetById(info.RequestListId);
                    var requestDetails = list.GetItemById(info.RequestListItemId);
                    pnpContext.Load(requestDetails);
                    pnpContext.ExecuteQuery();
                    // Using the value when not cleared
                    if (requestDetails["Owners"] != null)
                    {
                        //log.LogInformation("");
                        var userField = requestDetails.FieldValues["Owners"] as FieldUserValue[];
                        foreach (FieldUserValue userFieldValue in userField)
                        {
                            // Get the stored user lookup id value
                            var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", userFieldValue.Email);
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
                        var userField = requestDetails.FieldValues["Members"] as FieldUserValue[];
                        foreach (FieldUserValue userFieldValue in userField)
                        {
                            // Get the stored user lookup id value
                            var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", userFieldValue.Email);
                            var TeamUser = new AadUserConversationMember
                            {
                                Roles = new List<String>() { },
                                AdditionalData = new Dictionary<string, object>()
                               {{"user@odata.bind", objUser}}
                            };
                            Members.Add(TeamUser);
                        }
                    }
                    //Required Permissions:'TeamMember.ReadWrite.All'
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