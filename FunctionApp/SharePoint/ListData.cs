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

        /// <summary>
        /// Function to retrieve all companies.
        /// </summary>
        [FunctionName("GetCompanies")]
        public async Task<IActionResult> GetCompanies(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "all-companies")] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Companies list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.CompaniesListName);

            return new OkObjectResult(await list.GetListItems(new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title)"),
                new("orderby", "fields/Title")
            }));
        }

        /// <summary>
        /// Function to retrieve employees of a specific company.
        /// </summary>
        [FunctionName("GetEmployees")]
        public async Task<IActionResult> GetEmployees(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "all-companies/{companyId}/employees")] HttpRequest req,
            string company, ILogger log)
        {
            log.LogInformation("Employees list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.OperativeListName);

            return new OkObjectResult(await list.GetListItems(new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title,CSCS)"),
                new("filter", $"fields/Company eq '{company}'"),
                new("orderby", "fields/Title")
            }));
        }

        /// <summary>
        /// Function to retrieve signed-in companies for a specific site.
        /// </summary>
        [FunctionName("GetSignedInCompanies")]
        public async Task<IActionResult> GetSignedInCompanies(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sites/{siteName}/signed-in-companies")] HttpRequest req,
            string siteName, ILogger log)
        {
            log.LogInformation("Signed In companies list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.RegisterListName);

            var listItems = await list.GetListItems(new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Company)"),
                new("filter", $"fields/Site eq '{siteName}' and fields/CurrentStatus eq 'In'"),
                new("orderby", "fields/Company")
            });

            return new OkObjectResult(listItems.DistinctBy(i => i.Fields.AdditionalData["Company"]));
        }

        /// <summary>
        /// Function to retrieve signed-in employees for a specific company and site.
        /// </summary>
        [FunctionName("GetSignedInEmployees")]
        public async Task<IActionResult> GetSignedInEmployees(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "sites/{siteName}/companies/{companyId}/signed-in-employees")] HttpRequest req,
            string siteName, string company, ILogger log)
        {
            log.LogInformation("Signed In employees list is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.RegisterListName);

            var listItems = await list.GetListItems(new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "fields(select=Title)"),
                new("filter",
                    $"fields/Site eq '{siteName}' and fields/Company eq '{company}' and fields/CurrentStatus eq 'In'"),
                new("orderby", "fields/Title")
            });

            return new OkObjectResult(listItems.DistinctBy(i => i.Fields.AdditionalData["Title"]));
        }
    }
}
