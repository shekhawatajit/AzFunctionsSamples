using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using PnP.Framework.RER.Common.Tokens;
using System.Xml.Linq;
using PnP.Framework.RER.Common.Helpers;
using PnP.Framework.RER.Common.EventReceivers;
using System.Linq;
using Microsoft.SharePoint.Client;
using System.Net;

namespace PnP.Framework.RER.Functions
{
    public class ProjectRequestAdded
    {
        private readonly TokenManagerFactory _tokenManagerFactory;

        public ProjectRequestAdded(TokenManagerFactory tokenManagerFactory)
        {
            _tokenManagerFactory = tokenManagerFactory;
        }

        [FunctionName("ProjectRequestAdded")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            try
            {
                var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
                var xdoc = XDocument.Parse(requestBody);

                var eventRoot = xdoc.Root.Descendants().First().Descendants().First();

                if (eventRoot.Name.LocalName != "ProcessEvent" && eventRoot.Name.LocalName != "ProcessOneWayEvent")
                {
                    throw new Exception($"Unable to resolve event type");
                }

                var payload = eventRoot.FirstNode.ToString();
                var eventProperties = SerializerHelper.Deserialize<SPRemoteEventProperties>(payload);

                log.LogInformation(req.Host.Host);

                var host = req.Host.Host;

                var tokenManager = _tokenManagerFactory.Create(eventProperties, host);

                var context = await tokenManager.GetUserClientContextAsync(eventProperties.ItemEventProperties.WebUrl);
                context.Load(context.Web);
                await context.ExecuteQueryRetryAsync();

                if (eventRoot.Name.LocalName == "ProcessEvent")
                {
                    return await ProcessSyncEvent(eventProperties, context, log);
                }

                if (eventRoot.Name.LocalName == "ProcessOneWayEvent")
                {
                    return await ProcessAsyncEvent(eventProperties, context, log);
                }

                throw new Exception($"Unable to resolve event type");
            }
            catch (Exception ex)
            {
                log.LogError(new EventId(), ex, ex.Message);
                var result = new SPRemoteEventResult
                {
                    Status = SPRemoteEventServiceStatus.CancelWithError,
                    ErrorMessage = ex.Message
                };

                return new ContentResult
                {
                    Content = CreateEventResponse(result),
                    ContentType = "text/xml",
                    StatusCode = (int?)HttpStatusCode.InternalServerError
                };
            }
        }

        // -ing events, i.e ItemAdding
        private async Task<IActionResult> ProcessSyncEvent(SPRemoteEventProperties properties, ClientContext context, ILogger log)
        {
            switch (properties.EventType)
            {
                case SPRemoteEventType.ItemAdding:
                    {
                        context.Load(context.Web, p => p.Title);
                        context.ExecuteQueryRetry();
                        log.LogInformation(context.Web.Title);
                        break;
                    }
                //etc
                default: { break; }
            }
            var result = new SPRemoteEventResult
            {
                Status = SPRemoteEventServiceStatus.Continue
            };

            return new ContentResult
            {
                Content = CreateEventResponse(result),
                ContentType = "text/xml",
                StatusCode = (int?)HttpStatusCode.OK
            };
        }

        // -ed events, i.e. ItemAdded
        private async Task<IActionResult> ProcessAsyncEvent(SPRemoteEventProperties properties, ClientContext context, ILogger log)
        {
            switch (properties.EventType)
            {
                case SPRemoteEventType.ItemAdded:
                    {
                        // do things
                        break;
                    }
                //etc
                default: { break; }
            }
            return new OkResult();
        }

        private string CreateEventResponse(SPRemoteEventResult eventResult)
        {
            var responseTemplate = @"<s:Envelope xmlns:s=""http://schemas.xmlsoap.org/soap/envelope/"">
                                        <s:Body>{0}</s:Body>
                                    </s:Envelope>";
            var result = new ProcessEventResponse
            {
                ProcessEventResult = eventResult
            };
            var content = SerializerHelper.Serialize(result);

            return string.Format(responseTemplate, content);
        }
    }
}
