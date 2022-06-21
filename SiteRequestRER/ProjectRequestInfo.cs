using System;
using System.Collections.Generic;
namespace Onrocks.SharePoint
{
    public class ProjectRequestInfo
    {

        public int RequestListItemId { get; set; }
        public Guid RequestListId { get; set; }
        public string RequestSPSiteUrl { get; set; }
        public int RequestorId { get; set; }
    }
    public class FolderCreationInfo
    {
        public string LibraryName { get; set; }
        public List<string> Folders { get; set; }
    }
}