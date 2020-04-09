using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using MicrosoftGraphAspNetCoreConnectSample.Helpers;

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProxyController : ControllerBase
    {
        private readonly IGraphSdkHelper _graphSdkHelper;

        public ProxyController(IConfiguration configuration, IHostingEnvironment hostingEnvironment, IGraphSdkHelper graphSdkHelper)
        {
            _graphSdkHelper = graphSdkHelper;
        }

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
            var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);

            var qs = HttpContext.Request.QueryString;
            var url = $"{GetBaseUrlWithoutVersion(graphClient)}/{all}{qs.ToUriComponent()}";

            var request = new BaseRequest(url, graphClient, null)
            {
                Method = method,
                ContentType = HttpContext.Request.ContentType,
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
                using (var response = await request.SendRequestAsync(content?.ToString(), CancellationToken.None, HttpCompletionOption.ResponseContentRead).ConfigureAwait(false))
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

        private string GetBaseUrlWithoutVersion(GraphServiceClient graphClient)
        {
            var baseUrl = graphClient.BaseUrl;
            var index = baseUrl.LastIndexOf('/');
            return baseUrl.Substring(0, index);
        }
    }
}
