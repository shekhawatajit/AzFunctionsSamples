using System;
namespace Onrocks.SharePoint
{
    public class ProjectRequestInfo
    {

        public int ListItemId { get; set; }
        public Guid ListId { get; set; }
        public string WebUrl { get; set; }
        public int RequestorId { get; set; }
        public string? TeamsId { get; set; }
    }
}