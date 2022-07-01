using System;
namespace SimplifiedDelegatedRER
{
    public class AzureFunctionSettings
    {
#nullable enable
        public string? TenantId { get; set; }
        public string? ClientId { get; set; }
        public string? SecretName { get; set; }
        public string? KeyVaultName { get; set; }

        #region OIP
        public string? OIPHubSite { get; set; }
        public Guid OIPRequestListId { get; set; }
        public Guid OIPMailListId { get; set; }
        public string? OIPFolderInfojson { get; set; }
        public string? OIPProvisioningTemplateXmlFileUrl { get; set; }
        #endregion

        #region Relation
        public string? RelationHubSite { get; set; }
        public Guid RelationRequestListId { get; set; }
        public Guid RelationMailListId { get; set; }
        public string? RelationFolderInfojson { get; set; }
        public string? RelationProvisioningTemplateXmlFileUrl { get; set; }
        #endregion
#nullable disable
    }
}