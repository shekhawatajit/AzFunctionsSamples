using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Threading.Tasks;
using System.Text.Json;
using PnP.Core.Admin.Model.SharePoint;
using PnP.Framework;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using Microsoft.SharePoint.Client;
using System.Threading;
using System.IO;
using PnP.Core.QueryModel;
using System.Linq;
using System.Collections.Generic;

namespace Onrocks.SharePoint
{
    public class SetupSite
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly IPnPContextFactory pnpContextFactory;

        public SetupSite(IPnPContextFactory pnpContextFactory)
        {
            this.pnpContextFactory = pnpContextFactory;
        }
        [FunctionName("SetupSite")]
        [FixedDelayRetry(5, "00:00:10")]
        public async Task Run([QueueTrigger("%Step1QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
            log.LogInformation($"Setup of Sharepoint site process is started with data : {projectQueueItem}");
            try
            {
                string ProjectTitle, ProjectDescription, ProjectRequestor;
                string teamsSiteUrl;
                using (var contextPrimaryHub = await pnpContextFactory.CreateAsync(new Uri(info.RequestSPSiteUrl)))
                {
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
                    string templateUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, Environment.GetEnvironmentVariable("ProvisioningTemplateXmlFileUrl"));
                    IFile templateDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(templateUrl);
                    // Download the template file as stream
                    Stream downloadedContentStream = await templateDocument.GetContentAsync();
                    var provisioningTemplate = XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(downloadedContentStream);
                    log.LogInformation($"Template ID to apply :{provisioningTemplate.Id}");

                    //Reading Folder information
                    string folderInfoUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, Environment.GetEnvironmentVariable("FolderInfojson"));
                    IFile folderDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(folderInfoUrl);
                    // Download the template file as stream
                    Stream folderContentStream = await folderDocument.GetContentAsync();
                    StreamReader reader = new StreamReader(folderContentStream);
                    string folderjson = reader.ReadToEnd();
                    FolderCreationInfo folderInfo = JsonSerializer.Deserialize<FolderCreationInfo>(folderjson);

                    //Creating new request for Teams site without Group  
                    var teamsSiteToCreate = new TeamSiteWithoutGroupOptions(new Uri($"https://{contextPrimaryHub.Uri.DnsSafeHost}/sites/{uniqueSiteName}"), ProjectTitle)
                    {
                        Description = ProjectDescription,
                        //Language = Language.English,
                        Owner = $"i:0#.f|membership|{ProjectRequestor}"
                    };
                    teamsSiteUrl = teamsSiteToCreate.Url.AbsoluteUri;
                    log.LogInformation($"Creating site: {teamsSiteToCreate.Url}");

                    // Create the new site collection
                    using (var newSiteContext = await contextPrimaryHub.GetSiteCollectionManager().CreateSiteCollectionAsync(teamsSiteToCreate))
                    {
                        log.LogInformation($"Site created: {teamsSiteToCreate.Url}");
                        // Associating to Hub
                        ISite assocSite = await newSiteContext.Site.GetAsync(
                            p => p.HubSiteId,
                            p => p.IsHubSite);
                        if (assocSite.HubSiteId == Guid.Empty)
                        {
                            var resultJoin = await assocSite.JoinHubSiteAsync(primarySite.HubSiteId);
                            log.LogInformation($"Site connected to Hub: {resultJoin}");
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
                                    log.LogInformation(String.Format("{0:00}/{1:00} - {2}", progress, total, message));
                                },
                                IgnoreDuplicateDataRowErrors = true
                            };
                            web.ApplyProvisioningTemplate(provisioningTemplate, ptai);
                        }
                        // Creating Folders
                        CreateFolders(folderInfo, newSiteContext);
                    }
                    //Sending Email to Owner
                    UpdateSpList(ProjectTitle, ProjectDescription, ProjectRequestor, teamsSiteUrl, contextPrimaryHub);
                }
            }
            catch
            {
                throw;
            }
        }
        private async void CreateFolders(FolderCreationInfo folderInfo, PnPContext newSiteContext)
        {
            var folder = (newSiteContext.Web.Lists.GetByTitle(folderInfo.LibraryName, p => p.RootFolder)).RootFolder;

            foreach (string fld in folderInfo.Folders)
            {
                // Add a folder 
                var subFolder = await folder.EnsureFolderAsync(fld);
            }
        }
        private void UpdateSpList(string ProjectTitle, string ProjectDescription, string ProjectRequestor, string TeamSiteUrl, PnPContext contextPrimaryHub)
        {


            Guid MailListId = Guid.Parse(Environment.GetEnvironmentVariable("MailListId"));
            IList mailList = contextPrimaryHub.Web.Lists.GetById(MailListId, p => p.Title,
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
            var Receiver = contextPrimaryHub.Web.EnsureUser(ProjectRequestor);
            values.Add("Receiver", userfield.NewFieldUserValue(Receiver));
            var addedItem = mailList.Items.Add(values);
            addedItem.Update();
        }
    }
}