using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TranslateMyCode.Translate.com.microsofttranslator.api;

namespace TranslateMyCode.Translate
{
    public class BingTranslator : ITranslator
    {
        private string _clientID = null;
        private string _clientSecret = null; 
        private readonly string _translatorAccessURI = "https://datamarket.accesscontrol.windows.net/v2/OAuth2-13";
        private readonly string _requestDetails = "grant_type=client_credentials&client_id={0}&client_secret={1}&scope=http://api.microsofttranslator.com";
        private readonly string _translateBaseUrl = "http://api.microsofttranslator.com/v2/Http.svc/Translate?text=";

        public BingTranslator(string clientID, string clientSecret)
        {
            _clientID = clientID;
            _clientSecret = clientSecret;
        }
        // Unfortunately Bing doesn't offer a web api that could be used as a service discovery
        // but instead just offers API through basic POST requests. (2015-05-20)
        public string Translate(string word, Language from, Language to)
        {
            AdmAccessToken token = RequestAdmAccessToken();
            return RequestTranslation(word, token, from, to);           
        }

        private string RequestTranslation(string txtToTranslate, AdmAccessToken token, Language from, Language to)
        {
            // Create request for translation
            string url = _translateBaseUrl + System.Web.HttpUtility.UrlEncode(txtToTranslate) + FormatLanguageParameters(from, to);
            System.Net.WebRequest translationWebRequest = System.Net.WebRequest.Create(url);
            
            // Add authorization header with Adm token value
            string headerValue = "Bearer " + token.access_token;
            translationWebRequest.Headers.Add("Authorization", headerValue);

            // Fetch the response
            System.Net.WebResponse response = translationWebRequest.GetResponse();

            // XML format of given response by the BING API
            System.Xml.XmlDocument xTranslation = new System.Xml.XmlDocument();
            using (System.IO.Stream stream = response.GetResponseStream())
            {
                System.Text.Encoding encode = System.Text.Encoding.GetEncoding("utf-8");

                using (System.IO.StreamReader translatedStream = new System.IO.StreamReader(stream, encode))
                {
                    xTranslation.LoadXml(translatedStream.ReadToEnd());
                }
            }
            return xTranslation.InnerText;
        }

        /// <summary>
        /// Bing requires their Adm access token for translation
        /// </summary>
        /// <returns></returns>
        private AdmAccessToken RequestAdmAccessToken()
        {
            // Create the request for Adm access token
            System.Net.WebRequest webRequest = System.Net.WebRequest.Create(_translatorAccessURI);
            webRequest.ContentType = "application/x-www-form-urlencoded";
            webRequest.Method = "POST";

            // Create the contents for request
            var requestDetails = string.Format(_requestDetails, _clientID, _clientSecret);
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(requestDetails);
            webRequest.ContentLength = bytes.Length;

            // Write the contents into request stream
            using (System.IO.Stream outputStream = webRequest.GetRequestStream())
            {
                outputStream.Write(bytes, 0, bytes.Length);
            }

            // Fetch the response
            System.Net.WebResponse webResponse = webRequest.GetResponse();

            // Serialize and return the access token
            var serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(AdmAccessToken));
            return (AdmAccessToken)serializer.ReadObject(webResponse.GetResponseStream());
        }

        private string FormatLanguageParameters(Language from, Language to)
        {
            return string.Format("&from={0}&to={1}", from.ToString(), to.ToString());
        }
    }
}
