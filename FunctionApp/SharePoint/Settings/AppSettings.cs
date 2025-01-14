namespace Plumsail.DataSource.SharePoint.Settings
{
    public class AppSettings
    {
        public AzureApp AzureApp { get; set; }

        public ListData ListData { get; set; }
    }

    public class AzureApp
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string Tenant { get; set; }
    }

    public class ListData
    {
        public string SiteUrl { get; set; }
        public string MasterSiteURL { get; set; }
        public string CompaniesListName { get; set; }
        public string SiteListName { get; set; }
        public string OperativeListName { get; set; }
        public string RegisterListName { get; set; }
    }
}
