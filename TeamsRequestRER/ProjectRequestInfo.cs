using System;
namespace Adidas.OIP
{
    public class ProjectRequestInfo
    {

        public int RequestListItemId { get; set; }
        public Guid RequestListId { get; set; }
        public string RequestSPSiteUrl { get; set; }
        public int RequestorId { get; set; }
#nullable enable
        public string? TeamsId { get; set; }
#nullable disable
    }
}