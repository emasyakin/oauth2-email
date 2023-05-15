using System.Collections.Specialized;
using System.Web;

namespace oauth2_email.Helpers
{
    public static class UriHelper
    {
        public static UriBuilder AddToQuery(this UriBuilder uriBuilder, string name, string value)
        {
            var query = HttpUtility.ParseQueryString(uriBuilder.Query);
            query[name] = value;
            uriBuilder.Query = query.ToQueryString().TrimStart('?');
            return uriBuilder;
        }

        private static string ToQueryString(this NameValueCollection nvc)
        {
            var array = (from key in nvc.AllKeys
                    from value in nvc.GetValues(key)
                    select $"{HttpUtility.UrlEncode(key)}={HttpUtility.UrlEncode(value)}")
                .ToArray();
            return "?" + string.Join("&", array);
        }
    }
}
