using System;
using Azure.Security.KeyVault.Secrets;
using Azure.Identity;
using PnP.Core.Services;
using PnP.Core.QueryModel;
using System.Linq;
using System.Collections.Generic;
using PnP.Core.Model.SharePoint;
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
    }
}