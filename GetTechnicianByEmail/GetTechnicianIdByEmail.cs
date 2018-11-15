using System.Dynamic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace GetTechnicianByEmail
{
    public static class GetTechnicianIdByEmail
    {
        [FunctionName("GetTechnicianIdByEmail")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("FieldConnect HTTP trigger function processed a request.");

            // parse query parameter
            string email = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "email", true) == 0)
            .Value;

            var connectionString = @"AuthType=Office365;Url=https://<instance>.crm.dynamics.com/;Username=<username>;Password=<password>";
            CrmServiceClient conn = new CrmServiceClient(connectionString);

            IOrganizationService _orgService;
            _orgService = conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;

            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "cr4c8_email";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(email);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);

            QueryExpression query = new QueryExpression("cr4c8_fcotechnicians");
            query.ColumnSet.AddColumns("cr4c8_email");
            query.ColumnSet.AddColumns("cr4c8_techname");
            query.Criteria.AddFilter(filter1);

            EntityCollection result1 = _orgService.RetrieveMultiple(query);
            var result3 = result1.Entities.First().Attributes.Values;
            var technician = new { Email = result3.ElementAt(0), FullName = result3.ElementAt(1), Id = result3.ElementAt(2) };
            
            // Get request body
            dynamic data = await req.Content.ReadAsAsync<object>();

            string id = "";

            return id == null

            ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass an email on the query string or in the request body")

            : req.CreateResponse(HttpStatusCode.OK, "Technician ID: " + technician.Id.ToString());

        }
    }
}
