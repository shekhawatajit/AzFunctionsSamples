using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using Azure.Storage.Queues;
using System.IO;
using PnP.Core.Services;
using Microsoft.Graph;
namespace Onrocks.SharePoint
{
    public class SiteRequestAdded
    {
        private readonly IPnPContextFactory pnpContextFactory;
        public SiteRequestAdded(IPnPContextFactory pnpContextFactory)
        {
            this.pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("SiteRequestAdded")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("Item Added HTTP trigger function processed a request.");
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            string responseMessage = "This HTTP triggered function executed successfully.";
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(requestBody);

                string json = JsonConvert.SerializeXmlNode(xmlDoc);
                JObject eventData = JObject.Parse(json);

                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(new ProjectRequestInfo
                {
                    RequestListItemId = (int)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListItemId"],
                    RequestListId = Guid.Parse((string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListId"]),
                    RequestSPSiteUrl = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["WebUrl"],
                    RequestorId = (int)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["CurrentUserId"]
                });

                string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
                string QueueName = Environment.GetEnvironmentVariable("Step1QueueName");

                QueueClient theQueue = new QueueClient(connectionString, QueueName);
                var itemInfoBytes = System.Text.Encoding.UTF8.GetBytes(jsonString);
                await theQueue.SendMessageAsync(System.Convert.ToBase64String(itemInfoBytes));
            }
            catch (System.Exception err)
            {
                log.LogError(err.ToString());
                responseMessage = err.Message;
                return new BadRequestObjectResult(responseMessage);
            }
            return new OkObjectResult(responseMessage);
        }
    }
}