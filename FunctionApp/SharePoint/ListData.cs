using System.Collections.Generic;
using System.Linq;
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

        [FunctionName("GetCompanies")]
        public async Task<IActionResult> GetCompanies(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "all-companies")] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Companies list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.CompaniesListName);

            return new OkObjectResult(await GetListItems(list, new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title)"),
                new("orderby", "fields/Title")
            }));
        }

        [FunctionName("GetEmployees")]
        public async Task<IActionResult> GetEmployees(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "all-companies/{companyId}/employees")] HttpRequest req,
            string companyId, ILogger log)
        {
            log.LogInformation("Employees list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.OperativeListName);

            return new OkObjectResult(await GetListItems(list, new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title,CompanyLookupId,CSCS)"),
                new("filter", $"fields/CompanyLookupId eq '{companyId}'"),
                new("orderby", "fields/Title")
            }));
        }

        [FunctionName("GetSignedInCompanies")]
        public async Task<IActionResult> GetSignedInCompanies(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sites/{siteName}/signed-in-companies")] HttpRequest req,
            string siteName, ILogger log)
        {
            log.LogInformation("Signed In companies list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.RegisterListName);

            var listItems = await GetListItems(list, new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title,CurrentStatus,Company,CompanyLookupId,CSCS)"),
                new("filter", $"fields/Site eq '{siteName}' and fields/CurrentStatus eq 'In'"),
                new("orderby", "fields/Company")
            });

            return new OkObjectResult(listItems.DistinctBy(i => i.Fields.AdditionalData["Company"]));
        }

        [FunctionName("GetSignedInEmployees")]
        public async Task<IActionResult> GetSignedInEmployees(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sites/{siteName}/companies/{companyId}/signed-in-employees")] HttpRequest req,
            string siteName, string companyId, ILogger log)
        {
            log.LogInformation("Signed In employees list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.RegisterListName);

            return new OkObjectResult(await GetListItems(list, new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title,CurrentStatus,Company,CompanyLookupId,CSCS)"),
                new("filter", $"fields/Site eq '{siteName}' and fields/CompanyLookupId eq '{companyId}' and fields/CurrentStatus eq 'In'"),
                new("orderby", "fields/Title")
            }));
        }

        private static async Task<List<ListItem>> GetListItems(IListRequestBuilder list, List<QueryOption> queryOptions)
        {
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
