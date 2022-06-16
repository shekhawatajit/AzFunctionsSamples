using System;
using Microsoft.Azure.WebJobs;
using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using PnP.Framework;
using Microsoft.SharePoint.Client;
namespace Adidas.OIP
{
    public class CreateTeams
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly AzureFunctionSettings azureFunctionSettings;

        private readonly GraphServiceClient graphClient;
        public CreateTeams(AzureFunctionSettings settings, GraphServiceClient graphServiceClient)
        {
            this.azureFunctionSettings = settings;
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
                using (var pnpContext = new AuthenticationManager().GetACSAppOnlyContext(info.RequestSPSiteUrl, azureFunctionSettings.ClientId, azureFunctionSettings.ClientSecret))
                {
                    //Reading data from SharePoint list
                    GetProjectRequestDetails(info, out ProjectTitle, out ProjectDescription, out ProjectRequestor, pnpContext);

                    //Creating Teams (This is step 1/3)
                    string newTeamId = NewTeams(ProjectTitle, ProjectDescription, ProjectRequestor);

                    //Sending Teams info and request info in Queue 2
                    UpdateStep2Queue(info, newTeamId);

                    //Updating list item to send email to requestor
                    UpdateSpList(ProjectTitle, ProjectRequestor, pnpContext);
                }
            }
            catch (System.Exception err)
            {
                log.LogInformation(err.Message);
                log.LogInformation(err.StackTrace);
            }
        }

        private void UpdateSpList(string ProjectTitle, string ProjectRequestor, ClientContext pnpContext)
        {
            Guid MailListId = Guid.Parse(Environment.GetEnvironmentVariable("MailListId"));
            var mailList = pnpContext.Web.Lists.GetById(MailListId);

            //Item creation information
            var itemCreateInfo = new ListItemCreationInformation();
            var addedItem = mailList.AddItem(itemCreateInfo);
            addedItem["Title"] = $"Project Request {ProjectTitle}";
            addedItem["Status"] = "Request Accespted, Teams created";

            FieldUserValue[] users = new FieldUserValue[1];
            FieldUserValue Receiver = FieldUserValue.FromUser(ProjectRequestor);
            users[0] = Receiver;
            addedItem["Receiver"] = users;
            addedItem.Update();
            pnpContext.ExecuteQuery();
        }


        private string NewTeams(string ProjectTitle, string ProjectDescription, string ProjectRequestor)
        {
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

        private void GetProjectRequestDetails(ProjectRequestInfo info, out string ProjectTitle, out string ProjectDescription, out string ProjectRequestor, ClientContext pnpContext)
        {
            var list = pnpContext.Web.Lists.GetById(info.RequestListId);
            var requestDetails = list.GetItemById(info.RequestListItemId);
            ProjectTitle = requestDetails.FieldValues["Title"] == null ? string.Empty : requestDetails.FieldValues["Title"].ToString();
            ProjectDescription = requestDetails.FieldValues["Description"] == null ? string.Empty : requestDetails.FieldValues["Description"].ToString()!;
            ProjectRequestor = pnpContext.Web.GetUserById(info.RequestorId).UserPrincipalName;
        }
    }
}