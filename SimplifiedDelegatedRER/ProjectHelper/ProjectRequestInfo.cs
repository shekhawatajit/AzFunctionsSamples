using System;
using System.Collections.Generic;
using PnP.Core.Model.Security;

namespace SimplifiedDelegatedRER
{
    public class ProjectRequestInfo
    {
        public int RequestListItemId { get; set; }
        public int RequestorId { get; set; }
        public Guid RequestListId { get; set; }
#nullable enable
        public string? RequestSPSiteUrl { get; set; }
        public string? ProvisionTemplate { get; set; }
        public List<ISharePointUser>? Owners { get; set; }
        public List<ISharePointUser>? Members { get; set; }
        public List<ISharePointUser>? Visitors { get; set; }
        public string? ProjectTitle { get; set; }
        public string? ProjectDescription { get; set; }
        public string? ProjectRequestor { get; set; }
        public string? NewSiteUrl { get; set; }
        public string? SiteType { get; set; }
#nullable disable
    }
}