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
            bool ListSelector(List list)
            {
                return list.Name == listName || list.DisplayName == listName;
            }

            var url = new Uri(siteUrl);
            var queryOptions = new List<QueryOption>
            {
                new("select", "id"),
                new("expand", "lists(select=id,Name,DisplayName)")
            };

            var site = await graph.Sites.GetByPath(url.AbsolutePath, url.Host)
                .Request(queryOptions)
                .GetAsync();

            var listsPage = site.Lists;
            var list = listsPage.FirstOrDefault(ListSelector);

            while (list == null && listsPage.NextPageRequest != null)
            {
                listsPage = await listsPage.NextPageRequest.GetAsync();
                list = listsPage.FirstOrDefault(ListSelector);
            }

            return list != null
                ? graph.Sites[site.Id].Lists[list.Id]
                : null;
        }

        internal static async Task<List<ListItem>> GetListItems(this IListRequestBuilder list, List<QueryOption> queryOptions)
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
