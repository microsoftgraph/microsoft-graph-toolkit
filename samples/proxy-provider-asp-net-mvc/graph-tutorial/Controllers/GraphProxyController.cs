using graph_tutorial.Helpers;
using Microsoft.Graph;
using System;
using System.Collections;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace graph_tutorial.Controllers
{
    [RoutePrefix("api/GraphProxy")]
    public class GraphProxyController : ApiController
    {
        [HttpGet]
        [Route("{*all}")]
        public async Task<HttpResponseMessage> GetAsync(string all)
        {
            return await ProcessRequestAsync("GET", all, null).ConfigureAwait(false);
        }

        [HttpPost]
        [Route("{*all}")]
        public async Task<HttpResponseMessage> PostAsync(string all, [FromBody]object body)
        {
            return await ProcessRequestAsync("POST", all, body).ConfigureAwait(false);
        }

        [HttpDelete]
        [Route("{*all}")]
        public async Task<HttpResponseMessage> DeleteAsync(string all)
        {
            return await ProcessRequestAsync("DELETE", all, null).ConfigureAwait(false);
        }

        [HttpPut]
        [Route("{*all}")]
        public async Task<HttpResponseMessage> PutAsync(string all, [FromBody]object body)
        {
            return await ProcessRequestAsync("PUT", all, body).ConfigureAwait(false);
        }

        [HttpPatch]
        [Route("{*all}")]
        public async Task<HttpResponseMessage> PatchAsync(string all, [FromBody]object body)
        {
            return await ProcessRequestAsync("PATCH", all, body).ConfigureAwait(false);
        }

        private async Task<HttpResponseMessage> ProcessRequestAsync(string method, string all, object content)
        {
            var graphClient = GraphHelper.GetAuthenticatedClient();

            var request = new BaseRequest(GetURL(all, graphClient), graphClient, null)
            {
                Method = method,
                ContentType = HttpContext.Current.Request.ContentType,
            };

            var neededHeaders = Request.Headers.Where(h => h.Key.ToLower() == "if-match").ToList();
            if (neededHeaders.Count() > 0)
            {
                foreach (var header in neededHeaders)
                {
                    request.Headers.Add(new HeaderOption(header.Key, string.Join(",", header.Value)));
                }
            }
            

            var contentType = "application/json";

            try
            {
                using (var response = await request.SendRequestAsync(content, CancellationToken.None, HttpCompletionOption.ResponseContentRead).ConfigureAwait(false))
                {
                    response.Content.Headers.TryGetValues("content-type", out var contentTypes);

                    contentType = contentTypes?.FirstOrDefault() ?? contentType;

                    var byteArrayContent = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    return ReturnHttpResponseMessage(HttpStatusCode.OK, contentType, new ByteArrayContent(byteArrayContent));
                }
            }
            catch (ServiceException ex)
            {
                return ReturnHttpResponseMessage(ex.StatusCode, contentType, new StringContent(ex.Error.ToString()));
            }
        }

        private static HttpResponseMessage ReturnHttpResponseMessage(HttpStatusCode httpStatusCode, string contentType, HttpContent httpContent)
        {
            var httpResponseMessage = new HttpResponseMessage(httpStatusCode)
            {
                Content = httpContent
            };

            try
            {
                httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            }
            catch
            {
                httpResponseMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            }

            return httpResponseMessage;
        }

        private string GetURL(string all, GraphServiceClient graphClient)
        {
            var urlStringBuilder = new StringBuilder();

            var qs = HttpContext.Current.Request.QueryString;

            if (qs.Count > 0)
            {
                foreach (string key in qs.Keys)
                {
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    string[] values = qs.GetValues(key);
                    if (values == null) continue;

                    foreach (string value in values)
                    {
                        urlStringBuilder.Append(urlStringBuilder.Length == 0 ? "?" : "&");
                        urlStringBuilder.AppendFormat("{0}={1}", Uri.EscapeDataString(key), Uri.EscapeDataString(value));
                    }
                }
                urlStringBuilder.Insert(0, "?");

            }
            urlStringBuilder.Insert(0, $"{GetBaseUrlWithoutVersion(graphClient)}/{all}");

            return urlStringBuilder.ToString();
        }

        private string GetBaseUrlWithoutVersion(GraphServiceClient graphClient)
        {
            var baseUrl = graphClient.BaseUrl;
            var index = baseUrl.LastIndexOf('/');
            return baseUrl.Substring(0, index);
        }
    }
}
