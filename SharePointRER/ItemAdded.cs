using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Onrocks.SharePoint
{
    public static class ItemAdded
    {
        [FunctionName("ItemAdded")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("Item Added HTTP trigger function processed a request.");



            string xmldocument = "<?xml version='1.0' encoding='UTF-8'?> <s:Envelope xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'> <s:Body> <ProcessOneWayEvent xmlns='http://schemas.microsoft.com/sharepoint/remoteapp/'> <properties xmlns:i='http://www.w3.org/2001/XMLSchema-instance'> <AppEventProperties i:nil='true' /> <ContextToken /> <CorrelationId>25d640a0-e010-4000-22a5-cb49abba53a6</CorrelationId> <CultureLCID>1033</CultureLCID> <EntityInstanceEventProperties i:nil='true' /> <ErrorCode>GetContextTokenError</ErrorCode> <ErrorMessage>The app i:0i.t|ms.sp.ext|eb8ae9d1-a744-41da-9209-5ef678497bd6@51575b39-28de-4120-94c6-af4c743f70f1 does not have an endpoint or its endpoint is not valid.</ErrorMessage> <EventType>ItemAdded</EventType> <ItemEventProperties> <AfterProperties xmlns:a='http://schemas.microsoft.com/2003/10/Serialization/Arrays'> <a:KeyValueOfstringanyType> <a:Key>Owners</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>13;#i:0#.f|membership|ashutosh@onrocks.onmicrosoft.com</a:Value> </a:KeyValueOfstringanyType> <a:KeyValueOfstringanyType> <a:Key>ContentType</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>Item</a:Value> </a:KeyValueOfstringanyType> <a:KeyValueOfstringanyType> <a:Key>Members</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>11;#i:0#.f|membership|ajit@onrocks.onmicrosoft.com</a:Value> </a:KeyValueOfstringanyType> <a:KeyValueOfstringanyType> <a:Key>Title</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>Project A</a:Value> </a:KeyValueOfstringanyType> <a:KeyValueOfstringanyType> <a:Key>Visitors</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>18;#i:0#.f|membership|mandy@onrocks.onmicrosoft.com</a:Value> </a:KeyValueOfstringanyType> <a:KeyValueOfstringanyType> <a:Key>FileSystemObjectType</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>File</a:Value> </a:KeyValueOfstringanyType> <a:KeyValueOfstringanyType> <a:Key>TimesInUTC</a:Key> <a:Value xmlns:b='http://www.w3.org/2001/XMLSchema' i:type='b:string'>TRUE</a:Value> </a:KeyValueOfstringanyType> </AfterProperties> <AfterUrl i:nil='true' /> <BeforeProperties xmlns:a='http://schemas.microsoft.com/2003/10/Serialization/Arrays' /> <BeforeUrl /> <CurrentUserId>11</CurrentUserId> <ExternalNotificationMessage i:nil='true' /> <IsBackgroundSave>false</IsBackgroundSave> <ListId>35a72878-f5a9-42ad-bd56-9829b4079a33</ListId> <ListItemId>8</ListItemId> <ListTitle>Project Request</ListTitle> <UserDisplayName>Ajit Singh Shekhawat</UserDisplayName> <UserLoginName>i:0#.f|membership|ajit@onrocks.onmicrosoft.com</UserLoginName> <Versionless>false</Versionless> <WebUrl>https://onrocks.sharepoint.com/sites/RelationHubSite</WebUrl> </ItemEventProperties> <ListEventProperties i:nil='true' /> <SecurityEventProperties i:nil='true' /> <UICultureLCID>1033</UICultureLCID> <WebEventProperties i:nil='true' /> </properties> </ProcessOneWayEvent> </s:Body> </s:Envelope>";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmldocument);

            string json = JsonConvert.SerializeXmlNode(xmlDoc);
            JObject eventData = JObject.Parse(json);

            var ProjectTitle = eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["AfterProperties"]["a:KeyValueOfstringanyType"].ToList().Where(i => i["a:Key"].Value<string>() == "Title");
            var ProjectOwners = eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["AfterProperties"]["a:KeyValueOfstringanyType"].ToList().Where(i => i["a:Key"].Value<string>() == "Owners");
            var ProjectMembers = eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["AfterProperties"]["a:KeyValueOfstringanyType"].ToList().Where(i => i["a:Key"].Value<string>() == "Members");
            var ProjectVisitors = eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["AfterProperties"]["a:KeyValueOfstringanyType"].ToList().Where(i => i["a:Key"].Value<string>() == "Visitors");

            foreach (var pj in ProjectTitle)
            {
                log.LogInformation("Project Data: " + pj["a:Value"]["#text"].ToString().Split('|').Last());
            }
            foreach (var pj in ProjectOwners)
            {
                log.LogInformation("Project Data: " + pj["a:Value"]["#text"].ToString().Split('|').Last());
            }
            foreach (var pj in ProjectMembers)
            {
                log.LogInformation("Project Data: " + pj["a:Value"]["#text"].ToString().Split('|').Last());
            }
            foreach (var pj in ProjectVisitors)
            {
                log.LogInformation("Project Data: " + pj["a:Value"]["#text"].ToString().Split('|').Last());
            }
            string responseMessage = "This HTTP triggered function executed successfully.";
            return new OkObjectResult(responseMessage);
        }
    }
}
