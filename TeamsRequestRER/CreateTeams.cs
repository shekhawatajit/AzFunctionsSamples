using System;
using Microsoft.Azure.WebJobs;
using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;


namespace Onrocks.SharePoint
{
    public class CreateTeams
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly GraphServiceClient graphClient;
        public CreateTeams(IPnPContextFactory pnpContextFactory, GraphServiceClient graphServiceClient)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.graphClient = graphServiceClient;
        }

        [FunctionName("CreateTeams")]
        [FixedDelayRetry(5, "00:00:10")]
        public void Run([QueueTrigger("%Step1QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Team creation process started with data: {projectQueueItem}");
            string ProjectTitle, ProjectDescription, ProjectRequestor;
            try
            {
                using (var pnpContext = pnpContextFactory.Create(new Uri(info.RequestSPSiteUrl)))
                {
                    //Reading data from SharePoint list
                    GetProjectRequestDetails(info, out ProjectTitle, out ProjectDescription, out ProjectRequestor, pnpContext);

                    //Creating Teams (This is step 1/3)
                    string newTeamId = NewTeams(ProjectTitle, ProjectDescription, ProjectRequestor);

                    //Sending Teams info and request info in Queue 2
                    UpdateStep2Queue(info, newTeamId);

                    //Updating list item to send email to requestor
                   // UpdateSpList(ProjectTitle, ProjectRequestor, pnpContext);
                }
            }
            catch (System.Exception err)
            {
                log.LogInformation(err.Message);
                log.LogInformation(err.StackTrace);
                throw err;
            }
        }

    
        private string NewTeams(string ProjectTitle, string ProjectDescription, string ProjectRequestor)
        {
            //Required Permission: Microsoft Graph -> Team.Create
            var team = new Team
            {
                Visibility = TeamVisibilityType.Private,
                DisplayName = ProjectTitle,
                Description = ProjectDescription,
                AdditionalData = new Dictionary<string, object>() { { "template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')" } },
                Members = new TeamMembersCollectionPage()
                {
                    new AadUserConversationMember
                    {
                        Roles = new List<String>(){"owner"},
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('" + ProjectRequestor + "')"}
                        }
                    }
                },
            };
            var result = Task.Run(async () => await graphClient.Teams.Request().AddResponseAsync(team));
            string newTeamId = "";
            if (result.Result.HttpHeaders.TryGetValues("Location", out var locationValues))
            {
                newTeamId = locationValues?.First().Split('\'')[1];
            }
            return newTeamId;
        }
        private void UpdateStep2Queue(ProjectRequestInfo info, string newTeamId)
        {
            string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
            info.TeamsId = newTeamId;

            //Sending New teams Id in next queue for step 2
            var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(info);
            string QueueName = Environment.GetEnvironmentVariable("Step2QueueName");
            QueueClient theQueue = new QueueClient(connectionString, QueueName);
            var itemInfoBytes = System.Text.Encoding.UTF8.GetBytes(jsonString);
            theQueue.SendMessage(System.Convert.ToBase64String(itemInfoBytes));
        }

        private void GetProjectRequestDetails(ProjectRequestInfo info, out string ProjectTitle, out string ProjectDescription, out string ProjectRequestor, PnPContext pnpContext)
        {
            IList list = pnpContext.Web.Lists.GetById(info.RequestListId);
            IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                    li => li.Title,
                    li => li.All);
            ProjectTitle = requestDetails.Title == null ? string.Empty : requestDetails.Title;
            ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
            ProjectRequestor = pnpContext.Web.GetUserById(info.RequestorId).UserPrincipalName;
        }
    }
}