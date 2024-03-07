using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Plumsail.DataSource.SharePoint
{
    internal static class GraphServiceClientExtensions
    {
        internal static async Task<IListRequestBuilder> GetListAsync(this GraphServiceClient graph, string siteUrl, string listName)
        {
            var url = new Uri(siteUrl);
            var queryOptions = new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "lists(select=id,name)")
            };

            var site = await graph.Sites.GetByPath(url.AbsolutePath, url.Host)
                .Request(queryOptions)
                .GetAsync();

            var listsPage = site.Lists;
            var list = listsPage.FirstOrDefault(list => list.Name == listName);

            while (list == null && listsPage.NextPageRequest != null)
            {
                listsPage = await listsPage.NextPageRequest.GetAsync();
                list = listsPage.FirstOrDefault(l => l.Name == listName);
            }

            return list != null
                ? graph.Sites[site.Id].Lists[list.Id]
                : null;
        }
    }
}
