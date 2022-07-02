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
        public string? OIPRequestListTitle { get; set; }
        public string? OIPMailListTitle { get; set; }
        public string? OIPProvisioningTemplateXmlFileUrl { get; set; }
        #endregion

        #region Relation
        public string? RelationHubSite { get; set; }
        public string? RelationRequestListTitle { get; set; }
        public string? RelationMailListTitle { get; set; }
        public string? RelationProvisioningTemplateXmlFileUrl { get; set; }
        #endregion
#nullable disable
    }
}