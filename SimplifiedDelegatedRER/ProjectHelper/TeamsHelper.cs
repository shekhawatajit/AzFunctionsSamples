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


namespace SimplifiedDelegatedRER
{
    public class TeamsHelper
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly PnPContext _pnpContext;
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger _log;
        public TeamsHelper(PnPContext pnpContext, GraphServiceClient graphServiceClient, ILogger log)
        {
            this._pnpContext = pnpContext;
            this._graphClient = graphServiceClient;
            this._log = log;
        }

        public string CreateTeams(ProjectRequestInfo info)
        {

            _log.LogInformation("Team creation process started");
            string ProjectTitle, ProjectDescription, ProjectRequestor;
            try
            {

                //Reading data from SharePoint list
                GetProjectRequestDetails(info, out ProjectTitle, out ProjectDescription, out ProjectRequestor);

                //Creating Teams (This is step 1/3)
                return NewTeams(ProjectTitle, ProjectDescription, ProjectRequestor);
               
            }
            catch (System.Exception err)
            {
                _log.LogInformation(err.Message);
                _log.LogInformation(err.StackTrace);
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
            var result = Task.Run(async () => await _graphClient.Teams.Request().AddResponseAsync(team));
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

        private void GetProjectRequestDetails(ProjectRequestInfo info, out string ProjectTitle, out string ProjectDescription, out string ProjectRequestor)
        {
            IList list = _pnpContext.Web.Lists.GetById(info.RequestListId);
            IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                    li => li.Title,
                    li => li.All);
            ProjectTitle = requestDetails.Title == null ? string.Empty : requestDetails.Title;
            ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
            ProjectRequestor = _pnpContext.Web.GetUserById(info.RequestorId).UserPrincipalName;
        }
    }
}