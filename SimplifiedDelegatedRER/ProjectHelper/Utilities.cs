using System;
using Azure.Security.KeyVault.Secrets;
using Azure.Identity;
using PnP.Core.Services;
using PnP.Core.QueryModel;
using System.Linq;
using System.Collections.Generic;
using PnP.Core.Model.SharePoint;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using PnP.Core.Model.Security;
using Microsoft.Graph;
using PnP.Framework.Provisioning.Model;
using System.IO;
using PnP.Framework.Provisioning.Providers.Xml;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Threading;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace SimplifiedDelegatedRER
{
    public class Utilities
    {
        public string LoadSecret(string KeyVaultName, string SecretName)
        {
            var KeyVaultUrl = string.Format("https://{0}.vault.azure.net/", KeyVaultName);
            SecretClient client = new SecretClient(new Uri(KeyVaultUrl), new DefaultAzureCredential());
            return client.GetSecret(SecretName).Value.Value;
        }
        public void UpdateSpList(string MailListTitle, string ProjectTitle, string ProjectDescription, string ProjectRequestor, string TeamSiteUrl, PnPContext siteContext)
        {
            IList mailList = siteContext.Web.Lists.GetByTitle(MailListTitle, p => p.Title,
                                                                    p => p.Fields.QueryProperties(p => p.InternalName,
                                                                                                  p => p.FieldTypeKind,
                                                                                                  p => p.TypeAsString,
                                                                                                  p => p.Title));

            // Load Field
            IField userfield = mailList.Fields.Where(f => f.InternalName == "Receiver").FirstOrDefault()!;

            Dictionary<string, object> values = new Dictionary<string, object>
            {
                { "Title", $"Project Request {ProjectTitle}" },
                { "Status", $"Request Accespted, Teams created. SharePoint Site Url is {TeamSiteUrl}"}
            };
            var Receiver = siteContext.Web.EnsureUser(ProjectRequestor);
            values.Add("Receiver", userfield.NewFieldUserValue(Receiver));
            var addedItem = mailList.Items.Add(values);
            addedItem.Update();
        }
        public async Task<ProjectRequestInfo> ReadRequestFromList(PnPContext contextPrimaryHub, ProjectRequestInfo info, ILogger log)
        {
            log.LogInformation("ReadRequestFromList Started");
            //Reading data from SharePoint list and Geting the primary hub site details
            ISite primarySite = await contextPrimaryHub.Site.GetAsync(
                p => p.HubSiteId,
                p => p.IsHubSite);

            IList list = contextPrimaryHub.Web.Lists.GetByTitle(info.RequestListTitle);
            IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                    li => li.Title,
                    li => li.All);
            info.ProjectTitle = requestDetails.Title == null ? string.Empty : requestDetails.Title;
            info.ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
            info.ProjectRequestor = contextPrimaryHub.Web.GetUserById(info.RequestorId).UserPrincipalName;

            if (requestDetails["Owners"] != null)
            {
                info.Owners = new List<ISharePointUser>();
                foreach (IFieldUserValue user in (requestDetails["Owners"] as IFieldValueCollection)!.Values)
                {
                    // Get the stored user lookup id value
                    var spUser = contextPrimaryHub.Web.GetUserById(user.LookupId);
                    if (spUser != null)
                    {
                        info.Owners.Add(spUser);
                    }
                }
            }
            if (requestDetails["Members"] != null)
            {
                info.Members = new List<ISharePointUser>();
                foreach (IFieldUserValue user in (requestDetails["Members"] as IFieldValueCollection)!.Values)
                {
                    // Get the stored user lookup id value
                    var spUser = contextPrimaryHub.Web.GetUserById(user.LookupId);
                    if (spUser != null)
                    {
                        info.Members.Add(spUser);
                    }
                }
            }
            if (requestDetails["Visitors"] != null)
            {
                info.Visitors = new List<ISharePointUser>();
                foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                {
                    // Get the stored user lookup id value
                    var spUser = contextPrimaryHub.Web.GetUserById(user.LookupId);
                    if (spUser != null)
                    {
                        info.Visitors.Add(spUser);
                    }
                }
            }
            log.LogInformation("ReadRequestFromList Completed");
            return info;
        }
        public async Task CreateTeamsFromSPSite(PnPContext newTeamsSiteContext, ILogger log)
        {
            log.LogInformation("CreateTeamsFromSPSite Started");
            var site = await newTeamsSiteContext.Site.GetAsync(p => p.GroupId);
            Console.WriteLine($"Group ID is: {site.GroupId.ToString()}");
            //Creating Graph Client
            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
            {
                return newTeamsSiteContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com/v1.0"), requestMessage);
            }));
            var team = new Team
            {
                MemberSettings = new TeamMemberSettings
                {
                    AllowCreatePrivateChannels = true,
                    AllowCreateUpdateChannels = true
                },
                MessagingSettings = new TeamMessagingSettings
                {
                    AllowUserEditMessages = true,
                    AllowUserDeleteMessages = true
                },
                FunSettings = new TeamFunSettings
                {
                    AllowGiphy = true,
                    GiphyContentRating = GiphyRatingType.Strict
                }
            };
            await graphClient.Groups[site.GroupId.ToString()].Team.Request().PutAsync(team);
            log.LogInformation("CreateTeamsFromSPSite Complete");
        }
        public async Task<ProvisioningTemplate> ReadProvisionTemplte(PnPContext contextPrimaryHub, ILogger log, string TemplatePath)
        {
            log.LogInformation("ReadProvisionTemplte Started");
            string templateUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, TemplatePath);
            IFile templateDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(templateUrl);
            // Download the template file as stream
            Stream downloadedContentStream = await templateDocument.GetContentAsync();
            var provisioningTemplate = XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(downloadedContentStream);
            log.LogInformation($"Template ID :{provisioningTemplate.Id}");
            log.LogInformation("ReadProvisionTemplte Started");
            return provisioningTemplate;
        }
        public void ProvisionSite(PnPContext newTeamsSiteContext, ILogger log, ProvisioningTemplate template, ProjectRequestInfo info)
        {
            log.LogInformation("ProvisionSite Started");

            using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(newTeamsSiteContext))
            {
                // Use CSOM to load the web title
                csomContext.RequestTimeout = Timeout.Infinite;
                Web web = csomContext.Web;
                csomContext.Load(web, w => w.Title);
                csomContext.ExecuteQueryRetry();
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                {
                    ProgressDelegate = delegate (String message, Int32 progress, Int32 total)
                    {
                        Console.WriteLine(String.Format("{0:00}/{1:00} - {2}", progress, total, message));
                    },
                    IgnoreDuplicateDataRowErrors = true
                };
                web.ApplyProvisioningTemplate(template, ptai);
                log.LogInformation("ProvisionSite Completed");
            }
        }
        public async Task AddTeamMembers(PnPContext newTeamsSiteContext, ProjectRequestInfo info, ILogger log)
        {
            log.LogInformation("AddTeamMembers Started");

            var site = await newTeamsSiteContext.Site.GetAsync(p => p.GroupId);
            await newTeamsSiteContext.Web.LoadAsync(p => p.AssociatedVisitorGroup);
            Console.WriteLine($"Group ID is: {site.GroupId.ToString()}");
            //Creating Graph Client
            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
            {
                return newTeamsSiteContext.AuthenticationProvider.AuthenticateRequestAsync(new Uri("https://graph.microsoft.com/v1.0"), requestMessage);
            }));

            //Adding visitors directly to Site becuase Group does not support read only permission
            if (info.Visitors != null)
            {
                foreach (var user in info.Visitors)
                {
                    newTeamsSiteContext.Web.AssociatedVisitorGroup.AddUser(user.LoginName);
                }
            }

            //Adding members to team
            var Members = new List<ConversationMember>();
            if (info.Owners != null)
            {
                //log.LogInformation("");
                foreach (var user in info.Owners)
                {
                    // Get the stored user lookup id value
                    //requried Permission: Microsoft Graph -> User.Read.All
                    var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", user.UserPrincipalName);
                    var TeamUser = new AadUserConversationMember
                    {
                        Roles = new List<String>() { "owner" },
                        AdditionalData = new Dictionary<string, object>()
                               {{"user@odata.bind", objUser}}
                    };
                    Members.Add(TeamUser);
                }
            }
            if (info.Members != null)
            {
                foreach (var user in info.Members)
                {
                    // Get the stored user lookup id value
                    //requried Permission: Microsoft Graph -> User.Read.All
                    var objUser = string.Format("https://graph.microsoft.com/v1.0/users('{0}')", user.UserPrincipalName);
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
            var response = await graphClient.Teams[site.GroupId.ToString()].Members.Add(Members).Request().PostAsync();
            log.LogInformation("AddTeamMembers Completed");

        }
        public async Task AddSiteMembers(PnPContext newSiteContext, ProjectRequestInfo info, ILogger log)
        {
            log.LogInformation("AddSiteMembers Started");
            await newSiteContext.Web.LoadAsync(p => p.AssociatedOwnerGroup, p => p.AssociatedMemberGroup, p => p.AssociatedVisitorGroup);
           
            if (info.Owners != null)
            {
                foreach (var user in info.Owners)
                {
                    newSiteContext.Web.AssociatedOwnerGroup.AddUser(user.LoginName);
                }
            }
            if (info.Members != null)
            {
                foreach (var user in info.Members)
                {
                    newSiteContext.Web.AssociatedMemberGroup.AddUser(user.LoginName);
                }
            }
            //Adding visitors directly to Site becuase Group does not support read only permission
            if (info.Visitors != null)
            {
                foreach (var user in info.Visitors)
                {
                    newSiteContext.Web.AssociatedVisitorGroup.AddUser(user.LoginName);
                }
            }
            log.LogInformation("AddSiteMembers Completed");
        }
    }
}