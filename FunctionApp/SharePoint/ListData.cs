using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Plumsail.DataSource.SharePoint.Settings;

namespace Plumsail.DataSource.SharePoint
{
    public class ListData
    {
        private readonly Settings.ListData _settings;
        private readonly GraphServiceClientProvider _graphProvider;

        public ListData(IOptions<AppSettings> settings, GraphServiceClientProvider graphProvider)
        {
            _settings = settings.Value.ListData;
            _graphProvider = graphProvider;
        }

        [FunctionName("Companies")]
        public async Task<IActionResult> GetCompanies(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "companies")] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Companies list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.CompaniesListName);

            return new OkObjectResult(await GetListItems(list, new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title)"),
            }));
        }

        [FunctionName("Employees")]
        public async Task<IActionResult> GetEmployees(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "companies/{companyId}/employees")] HttpRequest req,
            string companyId, ILogger log)
        {
            log.LogInformation("Employees list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.EmployeesListName);

            return new OkObjectResult(await GetListItems(list, new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title,CompanyLookupId,CSCS)"),
                new("filter", $"fields/CompanyLookupId eq '{companyId}'")
            }));
        }

        private static async Task<List<ListItem>> GetListItems(IListRequestBuilder list, List<QueryOption> queryOptions)
        {
            var request = list.Items
                .Request(queryOptions);

            var itemsPage = await list.Items
                .Request(queryOptions)
                .GetAsync();

            var items = new List<ListItem>(itemsPage);

            while (itemsPage.NextPageRequest != null)
            {
                itemsPage = await itemsPage.NextPageRequest.GetAsync();
                items.AddRange(itemsPage);
            }

            return items;
        }
    }
}
