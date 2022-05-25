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
using System.Linq;
using System.Collections.Generic;
namespace Onrocks.SharePoint
{
    public static class ProjectRequestAdded
    {
        [FunctionName("ProjectRequestAdded")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
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

                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(new
                {
                    ListItemId = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListItemId"],
                    ListId = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["ListId"],
                    WebUrl = (string)eventData["s:Envelope"]["s:Body"]["ProcessOneWayEvent"]["properties"]["ItemEventProperties"]["WebUrl"]
                });

                string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
                string QueueName = Environment.GetEnvironmentVariable("QueueName");

                QueueClient theQueue = new QueueClient(connectionString, QueueName);
                await theQueue.SendMessageAsync(jsonString);
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
