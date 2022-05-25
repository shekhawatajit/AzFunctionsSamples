using System.Collections.Generic;
using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace OnRocks.SharePoint
{
    public class ProcessItemEvents
    {
        private readonly ILogger _logger;

        public ProcessItemEvents(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<ProcessItemEvents>();
        }

        [Function("ProcessItemEvents")]
        public HttpResponseData Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
        {
            _logger.LogInformation("C# HTTP trigger function processed a request.");

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");

            response.WriteString("Welcome to Azure Functions!");

            return response;
        }
    }
}
