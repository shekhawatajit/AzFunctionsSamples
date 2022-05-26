using System;
namespace Onrocks.SharePoint
{
    public class AzureFunctionSettings
    {
        #nullable enable
        public string? TenantId { get; set; }
        public string? ClientId { get; set; }
        public string? CertificateName { get; set; }
        public string? KeyVaultName { get; set; }
        public string? HubSite { get; set; }
        public Guid RequestListId { get; set; }
        public string? QueueName { get; set; }
        #nullable disable
    }
}