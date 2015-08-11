using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Web;
using System.Web.Helpers;
using JsonFx.Json;
using System.Dynamic;
using System.Text;

namespace GrokTaskPaneWeb.App
{
    public class AjaxProxy : IHttpHandler
    {

        #region IHttpHandler Members

        public bool IsReusable
        {
            // Return false in case your Managed Handler cannot be reused for another request.
            // Usually this would be false in case you have some state information preserved per request.
            get { return true; }
        }

        // example: https://github.com/numenta/grok-js-ua/blob/master/hello-grok/apiProxy.js
        public void ProcessRequest(HttpContext context)
        {
            dynamic proxyData;
            using (var inputReader = new StreamReader(context.Request.InputStream, context.Request.ContentEncoding))
            {
                //proxyData = Json.Decode(inputReader.ReadToEnd());
                var reader = new JsonFx.Json.JsonReader();
                proxyData = reader.Read(inputReader);
                //var serializer = new DataContractJsonSerializer(typeof(ProxyBody));
                //var proxyData = (ProxyBody)serializer.ReadObject(context.Request.InputStream);
            }

            string endpoint = proxyData.proxy.endpoint;
            bool hasData = ((System.Collections.Generic.IDictionary<String, Object>)proxyData.proxy).ContainsKey("data");

            // write the request data
            if (hasData && proxyData.proxy.method=="GET")
            {
                var ub = new UriBuilder(endpoint);
                var queryString = HttpUtility.ParseQueryString(ub.Query);
                foreach (var dataItem in ((System.Collections.Generic.IDictionary<String, Object>)proxyData.proxy.data))
                {
                    if (dataItem.Value == null) continue;
                    queryString.Add(dataItem.Key, dataItem.Value.ToString());
                }
                ub.Query = queryString.ToString();
                endpoint = ub.Uri.AbsoluteUri;
            }

            var request = WebRequest.CreateHttp(endpoint);
            request.Method = proxyData.proxy.method;

            request.Headers.Add(
                HttpRequestHeader.ContentEncoding,
                "UTF8"); 
            request.Headers.Add(
                HttpRequestHeader.Authorization, 
                GenerateBasicAuthenticationValue(username: proxyData.apiKey, password: string.Empty)); 

            try
            {
                // write the request data
                if (hasData && proxyData.proxy.method != "GET")
                {
                    using (var dataWriter = new StreamWriter(request.GetRequestStream(), new UTF8Encoding(false)))
                    {
                        var writer = new JsonFx.Json.JsonWriter();
                        writer.Write(proxyData.proxy.data, dataWriter);
                    }
                }

                // read the response 

                var response = (HttpWebResponse) request.GetResponse();
                context.Response.StatusCode = (int)HttpStatusCode.OK;
                context.Response.Headers.Add("Content-Encoding", response.ContentEncoding);

                byte[] buffer = new byte[ushort.MaxValue];
                using (var responseStream = response.GetResponseStream())
                {
                    int r;
                    while((r = responseStream.Read(buffer, 0, buffer.Length)) != 0)
                    {
                        context.Response.OutputStream.Write(buffer, 0, r);
                    }
                }

                return;
            }
            catch (WebException ex)
            {
                if (ex.Status != WebExceptionStatus.ProtocolError)
                {
                    context.Response.StatusCode = (int)HttpStatusCode.ServiceUnavailable;
                    return;
                }

                var httpResponse = (HttpWebResponse)ex.Response;
                context.Response.StatusCode = (int)httpResponse.StatusCode;
                return;
            }
        }

        private static string GenerateBasicAuthenticationValue(string username, string password)
        {
            var value = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(string.Format("{0}:{1}", username, password)));
            return string.Format("Basic {0}", value);
        }

        #endregion

        [Serializable]
        class ProxyBody
        {
            [DataMember(Order = 1)]
            public ProxyData proxy;

            [DataMember(Order = 2)]
            public string apiKey;
        }

        [Serializable]
        class ProxyData
        {
            [DataMember(Order = 1)]
            public string method;
            [DataMember(Order = 2)]
            public string endpoint;
        }
    }


   
}
