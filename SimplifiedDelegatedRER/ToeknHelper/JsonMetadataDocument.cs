using System.Collections.Generic;

namespace SimplifiedDelegatedRER.ToeknHelper
{
    public class JsonMetadataDocument
    {
        public List<JsonEndpoint> Endpoints { get; set; }
    }

    public class JsonEndpoint
    {
        public string Location { get; set; }
        public string Protocol { get; set; }
        public string Usage { get; set; }
    }
}