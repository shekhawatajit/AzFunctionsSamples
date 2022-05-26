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
namespace Onrocks.SharePoint
{
    public class ProjectRequestAdded
    {
        private readonly IPnPContextFactory pnpContextFactory;
        public ProjectRequestAdded(IPnPContextFactory pnpContextFactory)
        {
            this.pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("ProjectRequestAdded")]
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
                    ListItemId = (int)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListItemId"],
                    ListId = Guid.Parse((string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListId"]),
                    WebUrl = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["WebUrl"]
                });

                string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
                string QueueName = Environment.GetEnvironmentVariable("QueueName");

                QueueClient theQueue = new QueueClient(connectionString, QueueName);
                var itemInfoBytes = System.Text.Encoding.UTF8.GetBytes(jsonString);
                await theQueue.SendMessageAsync(System.Convert.ToBase64String(itemInfoBytes));
            }
            catch (System.Exception err)
            {
                log.LogError(err.Message);
                log.LogError(err.Source);
                responseMessage = err.Message;
            }
            return new OkObjectResult(responseMessage);
        }
    }
}
