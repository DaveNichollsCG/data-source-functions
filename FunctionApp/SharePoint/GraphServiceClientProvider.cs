using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Plumsail.DataSource.SharePoint.Settings;

namespace Plumsail.DataSource.SharePoint;

public class GraphServiceClientProvider
{
    private readonly AzureApp _azureAppSettings;

    public GraphServiceClientProvider(IOptions<AppSettings> settings)
    {
        _azureAppSettings = settings.Value.AzureApp;
    }

    public GraphServiceClient Create()
    {
        var confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(_azureAppSettings.ClientId)
            .WithClientSecret(_azureAppSettings.ClientSecret)
            .WithTenantId(_azureAppSettings.Tenant)
            .Build();

        var token = confidentialClientApplication
            .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
            .ExecuteAsync().Result.AccessToken;

        return new GraphServiceClient(new DelegateAuthenticationProvider(request =>
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            return Task.CompletedTask;
        }));
    }
}