using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using System.Text.Json;
using PnP.Core.Model.SharePoint;
using Microsoft.Graph;
using System.Collections.Generic;
using Azure.Storage.Queues;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;


namespace Onrocks.SharePoint
{
    public class CreateTeams
    {
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly GraphServiceClient graphClient;
        public CreateTeams(IPnPContextFactory pnpContextFactory, GraphServiceClient graphServiceClient)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.graphClient = graphServiceClient;
        }

        [FunctionName("CreateTeams")]
        public async void Run([QueueTrigger("%QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Queue trigger function processed: {projectQueueItem}");
            string ProjectTitle, ProjectDescription, ProjectRequestor;

            // Reading information from SharePoint list
            using (var pnpContext = pnpContextFactory.Create(new Uri(info.WebUrl)))
            {
                IListItem requestDetails = pnpContext.Web.Lists.GetById(info.ListId).Items.GetById(info.ListItemId);
                ProjectTitle = requestDetails.Title;
                ProjectDescription = requestDetails["Description"].ToString();
                ProjectRequestor = pnpContext.Web.GetUserById(info.RequestorId).UserPrincipalName;
            }
            //Creating Teams (This is step 1/3)
            var team = new Team
            {
                Visibility = TeamVisibilityType.Public,
                DisplayName = ProjectTitle,
                Description = ProjectDescription,
                AdditionalData = new Dictionary<string, object>() { { "template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')" } },
                Members = new TeamMembersCollectionPage()
                {
                    new AadUserConversationMember
                    {

                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"owners@odata.bind", "https://graph.microsoft.com/v1.0/users('" + ProjectRequestor + "')"}
                        }
                    }
                },
            };
            var newteam = graphClient.Teams.Request().AddAsync(team);

            string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
            info.TeamsId =  newteam.Result.Id;
            var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(info);

            string QueueName = Environment.GetEnvironmentVariable("Step2QueueName");
            QueueClient theQueue = new QueueClient(connectionString, QueueName);
            var itemInfoBytes = System.Text.Encoding.UTF8.GetBytes(jsonString);
            await theQueue.SendMessageAsync(System.Convert.ToBase64String(itemInfoBytes));
        }
    }
}
