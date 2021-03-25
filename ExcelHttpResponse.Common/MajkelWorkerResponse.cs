using System.Net;

namespace ExcelHttpResponse.Common
{
    public class MajkelWorkerResponse
    {
        public HttpStatusCode StatusCode { get; init; }
        public string RedirectionUrl { get; private set; }

        public void AddRedirectionUrl(string url)
            => (RedirectionUrl) = url;

        public override string ToString()
            => $"StatusCode: {StatusCode}{(string.IsNullOrEmpty(RedirectionUrl) == false ? $", RedirectionUrl: {RedirectionUrl}" : "")}";
    }
}
