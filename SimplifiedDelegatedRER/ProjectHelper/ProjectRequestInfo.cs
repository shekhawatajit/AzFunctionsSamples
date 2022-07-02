using System;
namespace SimplifiedDelegatedRER
{
    public class ProjectRequestInfo
    {
        public int RequestListItemId { get; set; }
        public Guid RequestListId { get; set; }
        public string RequestSPSiteUrl { get; set; }
        public int RequestorId { get; set; }
#nullable enable
        public string? TeamsId { get; set; }
        public string? ProjectTitle { get; set; }
        public string? ProjectDescription { get; set; }
        public string? ProjectRequestor { get; set; }
        public string? ProjectSite { get; set; }

#nullable disable
    }
}