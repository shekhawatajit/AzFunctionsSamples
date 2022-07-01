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
using PnP.Core.QueryModel;
using System.Linq;
using System.Collections.Generic;

namespace SimplifiedDelegatedRER
{
    public class TeamSiteHelper
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly PnPContext contextPrimaryHub;
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger _log;
        private readonly AzureFunctionSettings _settings;

        public TeamSiteHelper(PnPContext hubContext, GraphServiceClient graphServiceClient, ILogger log, AzureFunctionSettings functionSettings)
        {
            this.contextPrimaryHub = hubContext;
            this._graphClient = graphServiceClient;
            this._log = log;
            this._settings = functionSettings;
        }

        public async void SetupTeamSite(ProjectRequestInfo info, PnPContext pnpProjectSiteContext, Utilities ut)
        {
            //Get Teams site Url 
            //requried Permission: SharePoint -> Sites.FullControl.All
            var TeamSiteUrl = _graphClient.Groups[info.TeamsId].Sites["root"].Request().GetAsync().Result.WebUrl;
            _log.LogInformation($"Teams site Url to process : {TeamSiteUrl}");

            // Get the primary hub site details
            ISite primarySite = await contextPrimaryHub.Site.GetAsync(
                p => p.HubSiteId,
                p => p.IsHubSite);

            //Reading Visitors data
            IList list = contextPrimaryHub.Web.Lists.GetById(info.RequestListId);
            IListItem requestDetails = list.Items.GetById(info.RequestListItemId,
                    li => li.Title,
                    li => li.All);
            var ProjectTitle = requestDetails.Title == null ? string.Empty : requestDetails.Title;
            var ProjectDescription = requestDetails["Description"] == null ? string.Empty : requestDetails["Description"].ToString()!;
            var ProjectRequestor = contextPrimaryHub.Web.GetUserById(info.RequestorId).UserPrincipalName;


            //Reading Provisining Template
            string templateUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, _settings.OIPProvisioningTemplateXmlFileUrl);
            IFile templateDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(templateUrl);
            // Download the template file as stream
            Stream downloadedContentStream = await templateDocument.GetContentAsync();
            var provisioningTemplate = XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(downloadedContentStream);

            //Reading Folder information
            string folderInfoUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery,_settings.OIPFolderInfojson);
            IFile folderDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(folderInfoUrl);
            // Download the template file as stream
            Stream folderContentStream = await folderDocument.GetContentAsync();
            StreamReader reader = new StreamReader(folderContentStream);
            string folderjson = reader.ReadToEnd();
            FolderCreationInfo folderInfo = JsonSerializer.Deserialize<FolderCreationInfo>(folderjson);

            // Working on Teams Site                 // Associating to Hub
            ISite assocSite = await pnpProjectSiteContext.Site.GetAsync(
                p => p.HubSiteId,
                p => p.IsHubSite);
            if (assocSite.HubSiteId == Guid.Empty)
            {
                var resultJoin = await assocSite.JoinHubSiteAsync(primarySite.HubSiteId);
                _log.LogInformation($"Site connected to Hub: {resultJoin}");
            }

            //Adding visitors
            await pnpProjectSiteContext.Web.LoadAsync(p => p.AssociatedVisitorGroup);
            // Using the value when not cleared
            if (requestDetails["Visitors"] != null)
            {
                foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                {
                    // Get the stored user lookup id value
                    var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                    if (LoginName != null)
                    {
                        pnpProjectSiteContext.Web.AssociatedVisitorGroup.AddUser(LoginName);
                    }
                }
            }
            //Provisioning using PnP.Framework becuase PnP.Core does not work
            using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(pnpProjectSiteContext))
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
            // Creating Folders
            ut.CreateFolders(folderInfo, pnpProjectSiteContext);

            //Sending Email to Owner
            ut.UpdateSpList(_settings.OIPMailListId, ProjectTitle, ProjectDescription, ProjectRequestor, TeamSiteUrl, contextPrimaryHub);
        }
    }
}