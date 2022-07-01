using System;
using System.Collections.Generic;
namespace SimplifiedDelegatedRER
{
    public class RelationRequestInfo
    {
        public int RequestListItemId { get; set; }
        public Guid RequestListId { get; set; }
        public string RequestSPSiteUrl { get; set; }
        public int RequestorId { get; set; }
#nullable enable
        public string? RelationSiteUrl { get; set; }
        public string? SPSecret { get; set; }
        public string? ContextToken { get; set; }
#nullable disable
    }
}