using System;
using System.Collections.Generic;
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
        public string? SPSecret { get; set; }
        public string? ContextToken { get; set; }
#nullable disable
    }

    public class FolderCreationInfo
    {
        public string LibraryName { get; set; }
        public List<string> Folders { get; set; }
    }
}