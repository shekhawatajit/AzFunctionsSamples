using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Threading.Tasks;
using System.IO;
using PnP.Framework;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using Microsoft.SharePoint.Client;
using System.Threading;
using PnP.Core.QueryModel;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;

namespace Onrocks.SharePoint
{
    public class SetupProjectSite
    {
        // private readonly AzureFunctionSettings azureFunctionSettings;
        private readonly IPnPContextFactory pnpContextFactory;
        private readonly GraphServiceClient graphClient;
        public SetupProjectSite(IPnPContextFactory pnpContextFactory, GraphServiceClient graphServiceClient)
        {
            this.pnpContextFactory = pnpContextFactory;
            this.graphClient = graphServiceClient;
        }
        [FunctionName("SetupProjectSite")]
        [FixedDelayRetry(5, "00:00:10")]
        public async Task Run([QueueTrigger("%Step3QueueName%", Connection = "AzureWebJobsStorage")] string projectQueueItem, ILogger log)
        {
            try
            {
                ProjectRequestInfo info = System.Text.Json.JsonSerializer.Deserialize<ProjectRequestInfo>(projectQueueItem);
                log.LogInformation($"Setup of Sharepoint Teams site process is started with data : {projectQueueItem}");

                //Get Teams site Url 
                //requried Permission: SharePoint -> Sites.FullControl.All
                var TeamSiteUrl = graphClient.Groups[info.TeamsId].Sites["root"].Request().GetAsync().Result.WebUrl;
                log.LogInformation($"Teams site Url to process : {TeamSiteUrl}");
                using (var contextPrimaryHub = await pnpContextFactory.CreateAsync(new Uri(info.RequestSPSiteUrl)))
                {
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
                    string templateUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, Environment.GetEnvironmentVariable("ProvisioningTemplateXmlFileUrl"));
                    IFile templateDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(templateUrl);
                    // Download the template file as stream
                    Stream downloadedContentStream = await templateDocument.GetContentAsync();
                    var provisioningTemplate = XMLPnPSchemaFormatter.LatestFormatter.ToProvisioningTemplate(downloadedContentStream);

                    //Reading Folder information
                    string folderInfoUrl = string.Format("{0}{1}", contextPrimaryHub.Uri.PathAndQuery, Environment.GetEnvironmentVariable("FolderInfojson"));
                    IFile folderDocument = await contextPrimaryHub.Web.GetFileByServerRelativeUrlAsync(folderInfoUrl);
                    // Download the template file as stream
                    Stream folderContentStream = await folderDocument.GetContentAsync();
                    StreamReader reader = new StreamReader(folderContentStream);
                    string folderjson = reader.ReadToEnd();
                    FolderCreationInfo? folderInfo = JsonSerializer.Deserialize<FolderCreationInfo>(folderjson);

                    // Working on Teams Site
                    using (var context = await pnpContextFactory.CreateAsync(new Uri(TeamSiteUrl)))
                    {
                        // Associating to Hub
                        ISite assocSite = await context.Site.GetAsync(
                            p => p.HubSiteId,
                            p => p.IsHubSite);
                        if (assocSite.HubSiteId == Guid.Empty)
                        {
                            var resultJoin = await assocSite.JoinHubSiteAsync(primarySite.HubSiteId);
                            log.LogInformation($"Site connected to Hub: {resultJoin}");
                        }

                        //Adding visitors
                        await context.Web.LoadAsync(p => p.AssociatedVisitorGroup);
                        // Using the value when not cleared
                        if (requestDetails["Visitors"] != null)
                        {
                            foreach (IFieldUserValue user in (requestDetails["Visitors"] as IFieldValueCollection)!.Values)
                            {
                                // Get the stored user lookup id value
                                var LoginName = contextPrimaryHub.Web.GetUserById(user.LookupId).LoginName;
                                if (LoginName != null)
                                {
                                    context.Web.AssociatedVisitorGroup.AddUser(LoginName);
                                }
                            }
                        }
                        //Provisioning using PnP.Framework becuase PnP.Core does not work
                        using (ClientContext csomContext = PnPCoreSdk.Instance.GetClientContext(context))
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
                        CreateFolders(folderInfo, context);
                    }
                    //Sending Email to Owner
                    UpdateSpList(ProjectTitle, ProjectDescription, ProjectRequestor, TeamSiteUrl, contextPrimaryHub);
                }
            }
            catch
            {
                throw;
            }
        }

        private void CreateFolders(FolderCreationInfo? folderInfo, PnPContext NewSiteContext)
        {
            var folder = (NewSiteContext.Web.Lists.GetByTitle(folderInfo.LibraryName, p => p.RootFolder)).RootFolder;

            foreach (string fld in folderInfo.Folders)
            {
                // Add a folder 
                var subFolder = folder.Folders.Add(fld);
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