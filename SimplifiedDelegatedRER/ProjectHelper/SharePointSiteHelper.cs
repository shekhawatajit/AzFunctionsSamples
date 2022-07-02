using System;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.IO;
using PnP.Framework;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System.Text.Json;
using Microsoft.SharePoint.Client;
using System.Threading;
using PnP.Core.Admin.Model.SharePoint;
using System.Threading.Tasks;
namespace SimplifiedDelegatedRER
{
    public class SharePointSiteHelper
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly PnPContext contextPrimaryHub;
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger _log;
        private readonly AzureFunctionSettings _settings;
        public SharePointSiteHelper(PnPContext hubContext, GraphServiceClient graphServiceClient, ILogger log, AzureFunctionSettings functionSettings)
        {
            this.contextPrimaryHub = hubContext;
            this._graphClient = graphServiceClient;
            this._log = log;
            this._settings = functionSettings;
        }
        public async Task<string> SetupSite(RelationRequestInfo info, Utilities ut)
        {
            _log.LogInformation($"Setup of Sharepoint site process is started with data : {info}");
            try
            {
                string ProjectTitle, ProjectDescription, ProjectRequestor, teamsSiteUrl;

                //Reading data from SharePoint list and Geting the primary hub site details
                ISite primarySite = await contextPrimaryHub.Site.GetAsync(
                    p => p.HubSiteId,
                    p => p.IsHubSite);

                IList list = contextPrimaryHub.Web.Lists.GetById(info.RequestListId);
                IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                        li => li.Title,
                        li => li.All);
                ProjectTitle = requestDetails.Title == null ? string.Empty : requestDetails.Title;
                ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
                ProjectRequestor = contextPrimaryHub.Web.GetUserById(info.RequestorId).UserPrincipalName;

                //Generating Unique site Url
                string uniqueSiteName = Guid.NewGuid().ToString().Split('-')[4];

                //Reading Provisining Template
                string templateUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, _settings.RelationProvisioningTemplateXmlFileUrl);
                IFile templateDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(templateUrl);
                // Download the template file as stream
                Stream downloadedContentStream = await templateDocument.GetContentAsync();
                var provisioningTemplate = XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(downloadedContentStream);
                _log.LogInformation($"Template ID to apply :{provisioningTemplate.Id}");

                //Creating new request for Teams site without Group  
                var teamsSiteToCreate = new TeamSiteWithoutGroupOptions(new Uri($"https://{contextPrimaryHub.Uri.DnsSafeHost}/sites/{uniqueSiteName}"), ProjectTitle)
                {
                    Description = ProjectDescription,
                    //Language = Language.English,
                    Owner = $"i:0#.f|membership|{ProjectRequestor}"
                };
                teamsSiteUrl = teamsSiteToCreate.Url.AbsoluteUri;
                _log.LogInformation($"Creating site: {teamsSiteToCreate.Url}");

                // Create the new site collection
                using (var newSiteContext = await contextPrimaryHub.GetSiteCollectionManager().CreateSiteCollectionAsync(teamsSiteToCreate))
                {
                    _log.LogInformation($"Site created: {teamsSiteToCreate.Url}");
                    // Associating to Hub
                    ISite assocSite = await newSiteContext.Site.GetAsync(
                        p => p.HubSiteId,
                        p => p.IsHubSite);
                    if (assocSite.HubSiteId == Guid.Empty)
                    {
                        var resultJoin = await assocSite.JoinHubSiteAsync(primarySite.HubSiteId);
                        _log.LogInformation($"Site connected to Hub: {resultJoin}");
                    }

                    //Adding Users
                    await newSiteContext.Web.LoadAsync(p => p.AssociatedOwnerGroup, p => p.AssociatedMemberGroup, p => p.AssociatedVisitorGroup);
                    // Using the value when not cleared
                    if (requestDetails["Owners"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Owners"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                            if (LoginName != null)
                            {
                                newSiteContext.Web.AssociatedOwnerGroup.AddUser(LoginName);
                            }
                        }
                    }
                    if (requestDetails["Members"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Members"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                            if (LoginName != null)
                            {
                                newSiteContext.Web.AssociatedMemberGroup.AddUser(LoginName);
                            }
                        }
                    }
                    if (requestDetails["Visitors"] != null)
                    {
                        foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                        {
                            // Get the stored user lookup id value
                            var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                            if (LoginName != null)
                            {
                                newSiteContext.Web.AssociatedVisitorGroup.AddUser(LoginName);
                            }
                        }
                    }

                    //Provisioning using PnP.Framework becuase PnP.Core does not work
                    using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(newSiteContext))
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
                                _log.LogInformation(String.Format("{0:00}/{1:00} - {2}", progress, total, message));
                            },
                            IgnoreDuplicateDataRowErrors = true
                        };
                        web.ApplyProvisioningTemplate(provisioningTemplate, ptai);
                    }
                }
                //Sending Email to Owner
                ut.UpdateSpList(_settings.RelationMailListTitle, ProjectTitle, ProjectDescription, ProjectRequestor, teamsSiteUrl, contextPrimaryHub);
                return teamsSiteUrl;
            }
            catch
            {
                throw;
            }
        }
    }
}