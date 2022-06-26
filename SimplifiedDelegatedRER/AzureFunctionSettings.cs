using System;
namespace  SimplifiedDelegatedRER
{
    public class AzureFunctionSettings
    {
#nullable enable
        public string? TenantId { get; set; }
        public string? ClientId { get; set; }
        public string? SPClientId { get; set; }
        public string? SecretName { get; set; }
        public string? SPSecretName { get; set; }
        public string? KeyVaultName { get; set; }
        public string? HubSite { get; set; }
        public Guid RequestListId { get; set; }
        public Guid MailListId { get; set; }
#nullable disable
    }
}