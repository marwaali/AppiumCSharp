using log4net;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Cookie = System.Net.Cookie;

namespace UIAutomation.Utils_Misc
{
    public class WebService : BasePage
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(WebService));

        private string xRequestToken;
        private string APIKey;
        private string APIUrl;
        private string token;
        private string encodedToken;
        private CookieContainer LEGACY_SITE_COOKIES = null;

        public WebService()
        {
            APIKey = Data.Get("config:apikey");
            APIUrl = Data.Get("config:apiurl");
            token = GetToken();
            encodedToken = Convert.ToBase64String(Encoding.UTF8.GetBytes(token));
        }

        private string GetToken()
        {
            string postString = $"{{ApiKey:'{APIKey}',UserName:'{Data.Get("config:username")}',Password:'{Data.Get("config:password")}',Version:1}}";
            string url = $"{APIUrl}login";
            HttpContent postContent = new StringContent(postString, Encoding.UTF8, "application/json");
            string responseString;
            var baseAddress = new Uri(url);
            using (var client = new HttpClient {BaseAddress = baseAddress})
            {
                HttpResponseMessage result = client.PostAsync("", postContent).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return (string) JObject.Parse(responseString).SelectToken("$.Token");
        }

        public string GetApiRequest(string url, string parameters = null)
        {
            string responseString;
            var baseAddress = string.IsNullOrEmpty(parameters) ? new Uri(APIUrl + url) : new Uri($"{APIUrl}{url}?{parameters}");
            using (var client = new HttpClient {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Add("authorization", $"Basic {encodedToken}");
                var result = client.GetAsync(baseAddress).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }

        public string DeleteApiRequest(string url)
        {
            string responseString;
            var baseAddress = new Uri($"{APIUrl}{url}");
            using (var client = new HttpClient {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Add("authorization", $"Basic {encodedToken}");
                var result = client.DeleteAsync(baseAddress).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }

        public string PostApiRequest(string postUrl, string postString, string postContentType = "application/json")
        {
            HttpContent postContent = new StringContent(postString, Encoding.UTF8, postContentType);
            string responseString;
            var baseAddress = new Uri(APIUrl + postUrl);
            using (var client = new HttpClient {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(postContentType));
                client.DefaultRequestHeaders.Add("authorization", $"Basic {encodedToken}");
                HttpResponseMessage result = client.PostAsync("", postContent).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }

        public string PostRequest(string postUrl, string postString, string postContentType = "application/json")
        {
            HttpContent postContent = new StringContent(postString, Encoding.UTF8, postContentType);
            string responseString;
            var baseAddress = new Uri(GetBaseUrl() + postUrl);
            using (var handler = new HttpClientHandler {CookieContainer = GetCookies()})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(postContentType));
                client.DefaultRequestHeaders.Add("X-Request-Verification-Token", xRequestToken);
                client.DefaultRequestHeaders.Add("Accept-Language", "en-US,ru-RU;q=0.8,ru;q=0.5,en;q=0.3");
                HttpResponseMessage result = client.PostAsync("", postContent).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }

        //Wrapper for "PostFormRequest" which accepts Dictionary and converts it to the list
        public string PostFormRequest(string postUrl, Dictionary<string, string> formData, string postContentType = "*/*")
        {
            List<KeyValuePair<string, string>> form = formData.Select(x => new KeyValuePair<string, string>(x.Key, x.Value)).ToList();
            return PostFormRequest(postUrl, form, postContentType);
        }

        //Performs a post request with form data. Accepts data form as List<KeyValuePair<string, string>>
        public string PostFormRequest(string postUrl, List<KeyValuePair<string, string>> formData, string postContentType = "*/*", CookieContainer cookies = null)
        {
            string responseString;
            cookies = cookies ?? GetCookies();
            string fullUrl = postUrl.Contains("http") ? postUrl : GetBaseUrl() + postUrl;
            var baseAddress = new Uri(fullUrl);

            var encodedItems = formData.Select(i => $"{WebUtility.UrlEncode(i.Key)}={WebUtility.UrlEncode(i.Value)}");
            var encodedContent = new StringContent(String.Join("&", encodedItems), null, "application/x-www-form-urlencoded");

            var req = new HttpRequestMessage(HttpMethod.Post, fullUrl) {Content = encodedContent};
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(postContentType));
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(postContentType));
                client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.5");
                client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0");
                HttpResponseMessage result = client.SendAsync(req).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }

        // Performs a post request to download file
        public string PostFormDownloadFileRequest(string postUrl, List<KeyValuePair<string, string>> formData, string postContentType = "*/*", CookieContainer cookies = null)
        {
            var folderLocation = Data.Get("config:fileDownloadPath");
            string generatedName;
            cookies = cookies ?? GetCookies();
            string fullUrl = postUrl.Contains("http") ? postUrl : GetBaseUrl() + postUrl;
            var baseAddress = new Uri(fullUrl);

            var encodedItems = formData.Select(i => $"{WebUtility.UrlEncode(i.Key)}={WebUtility.UrlEncode(i.Value)}");
            var encodedContent = new StringContent(String.Join("&", encodedItems), null, "application/x-www-form-urlencoded");

            var req = new HttpRequestMessage(HttpMethod.Post, fullUrl) {Content = encodedContent};
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(postContentType));
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(postContentType));
                client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.5");
                client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0");
                HttpResponseMessage result = client.SendAsync(req).Result;
                result.EnsureSuccessStatusCode();
                var fileName = result.Content.Headers.ContentDisposition.FileName.Replace("\"", "");
                generatedName = folderLocation + fileName;
                var response = result.Content.ReadAsStreamAsync();
                using (var fileStream = File.Create(generatedName))
                using (var reader = new StreamReader(response.Result))
                {
                    response.Result.CopyTo(fileStream);
                    fileStream.Flush();
                }
            }
            return generatedName;
        }

        //Uploads file
        public string UploadFile(string postUrl, Dictionary<string, string> form, string filePath, CookieContainer cookies = null)
        {
            cookies = cookies ?? GetCookies();
            string fullUrl = postUrl.Contains("http") ? postUrl : GetBaseUrl() + postUrl;

            // Read file data
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            byte[] data = new byte[fs.Length];
            fs.Read(data, 0, data.Length);
            fs.Close();

            Dictionary<string, object> formObject = new Dictionary<string, object>();
            foreach (var f in form)
            {
                formObject.Add(f.Key, f.Value);
            }

            string fileName = filePath.Split('\\').Last();
            string contentType = GetContentTypeByFileExtension(fileName);
            formObject["upl"] = new FormUpload.FileParameter(data, fileName, contentType);

            // Create request and receive response
            string userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0";
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(fullUrl, userAgent, formObject, cookies);

            // Process response
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            webResponse.Close();
            return fullResponse;
        }

        //Used to upload several files or fake form data with multipart-form-data object
        public string UploadFiles(string postUrl, Dictionary<string, string> form, List<FormUpload.FileParameter> files, CookieContainer cookies)
        {
            string fullUrl = postUrl.Contains("http") ? postUrl : GetBaseUrl() + postUrl;
            cookies = cookies ?? GetCookies();

            Dictionary<string, object> formObject = new Dictionary<string, object>();
            foreach (var f in form)
            {
                formObject.Add(f.Key, f.Value);
            }

            foreach (var file in files)
            {
                formObject[file.FormParamName] = file;
            }

            // Create request and receive response
            string userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0";
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(fullUrl, userAgent, formObject, cookies);

            // Process response
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            webResponse.Close();
            return fullResponse;
        }

        public string GetRequest(string url, CookieContainer cookies = null)
        {
            string responseString;
            cookies = cookies ?? GetCookies();
            var fullUrl = url.Contains("http") ? url : GetBaseUrl() + url;
            var baseAddress = new Uri(fullUrl);
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0");
                client.DefaultRequestHeaders.ConnectionClose = false;
                client.DefaultRequestHeaders.Host = fullUrl.Split('/')[2];
                var result = client.GetAsync(baseAddress).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }

        public string OptionsRequest(string postUrl, CookieContainer cookies = null)
        {
            string responseString;
            cookies = cookies ?? GetCookies();
            string fullUrl = postUrl.Contains("http") ? postUrl : GetBaseUrl() + postUrl;
            var baseAddress = new Uri(fullUrl);

            var req = new HttpRequestMessage(HttpMethod.Options, fullUrl);
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
                client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.5");
                client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0");
                HttpResponseMessage result = client.SendAsync(req).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return responseString;
        }
        
        
        public string DownloadFileByLinkWithCookies(string url, string folderLocation = null, CookieContainer cookies = null)
        {
            cookies = cookies ?? GetCookies();
            string fullUrl = url.Contains("http") ? url : GetBaseUrl() + url;
            var baseAddress = new Uri(fullUrl);
            folderLocation = folderLocation ?? Data.Get("config:filedownloadpath");
            string generatedName;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler))
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0");
                client.DefaultRequestHeaders.ConnectionClose = false;
                var result = client.GetAsync(fullUrl).Result;
                result.EnsureSuccessStatusCode();
                string fileName = result.Content.Headers.ContentDisposition.FileName.Replace("\"", "");
                generatedName = folderLocation + fileName;
                if (new FileInfo(generatedName).Exists)
                {
                    new FileInfo(generatedName).Delete();
                }
                var response = result.Content.ReadAsStreamAsync();
                using (var fileStream = File.Create(generatedName))
                using (var reader = new StreamReader(response.Result))
                {
                    response.Result.CopyTo(fileStream);
                    fileStream.Flush();
                }
            }
            return generatedName;
        }

        public string DownloadFileByLinkWithoutCookiesAndHosts(string url, string folderLocation = null)
        {
            string fullUrl = url.Contains("http") ? url : GetBaseUrl() + url;
            folderLocation = folderLocation ?? Data.Get("config:filedownloadpath");
            string generatedName;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (var handler = new HttpClientHandler())
            using (var client = new HttpClient(handler))
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0");
                client.DefaultRequestHeaders.ConnectionClose = false;
                var result = client.GetAsync(fullUrl).Result;
                result.EnsureSuccessStatusCode();
                string fileName = result.Content.Headers.ContentDisposition.FileName.Replace("\"", "");
                generatedName = folderLocation + fileName;
                if (new FileInfo(generatedName).Exists)
                {
                    new FileInfo(generatedName).Delete();
                }
                var response = result.Content.ReadAsStreamAsync();
                using (var fileStream = File.Create(generatedName))
                using (var reader = new StreamReader(response.Result))
                {
                    response.Result.CopyTo(fileStream);
                    fileStream.Flush();
                }
            }
            return generatedName;
        }

        private CookieContainer GetCookies()
        {
            string handle = null;
            ICookieJar c = driver.Manage().Cookies;
            // added for IE driver
            if (!c.AllCookies.Any())
            {
                handle = driver.CurrentWindowHandle;
                driver.SwitchTo().Window(driver.WindowHandles[0]);
                c = driver.Manage().Cookies;
            }
            CookieContainer cookieContainer = new CookieContainer();
            foreach (var t in c.AllCookies)
            {
                cookieContainer.Add(new Cookie(t.Name, t.Value) {Domain = (new Uri(GetBaseUrl())).Host});
            }
            if (handle != null)
            {
                driver.SwitchTo().Window(handle);
            }
            return cookieContainer;
        }


        // HERE GOES METHODS TO WORK WITH LEGACY REQUESTS WITHOUT BROWSER SESSION

        // Get request for legacy site without using webdriver and browser. Used WS authorization
        public string Legacy_GetRequest(string url)
        {
            string fullUrl = url.Contains("http") ? url : _getLegacyBaseAddress() + "/" + url;
            return GetRequest(fullUrl, _getLegacyCookies());
        }

        //Wrapper for "Legacy_PostFormRequest" which accepts Dictionary and converts it to the list
        public string Legacy_PostFormRequest(string postUrl, Dictionary<string, string> formData, string postContentType = "*/*")
        {
            List<KeyValuePair<string, string>> form = formData.Select(x => new KeyValuePair<string, string>(x.Key, x.Value)).ToList();
            return Legacy_PostFormRequest(postUrl, form, postContentType);
        }

        // PostForm request for legacy site without using webdriver and browser. Used WS authorization
        public string Legacy_PostFormRequest(string url, List<KeyValuePair<string, string>> formData, string postContentType = "*/*")
        {
            string fullUrl = url.Contains("http") ? url : _getLegacyBaseAddress() + "/" + url;
            return PostFormRequest(fullUrl, formData, postContentType, _getLegacyCookies());
        }
        
        // PostForm request for download file
        public string Legacy_PostFormDownloadFileRequest(string url, Dictionary<string, string> formData, string postContentType = "*/*")
        {
            List<KeyValuePair<string, string>> form = formData.Select(x => new KeyValuePair<string, string>(x.Key, x.Value)).ToList();
            return Legacy_PostFormDownloadFileRequest(url, form, postContentType);
        }
        
        // PostForm request for download file
        public string Legacy_PostFormDownloadFileRequest(string url, List<KeyValuePair<string, string>> formData, string postContentType = "*/*")
        {
            string fullUrl = url.Contains("http") ? url : _getLegacyBaseAddress() + "/" + url;
            return PostFormDownloadFileRequest(fullUrl, formData, postContentType, _getLegacyCookies());
        }

        //Uploads file
        public string Legacy_UploadFile(string postUrl, Dictionary<string, string> form, string filePath)
        {
            string fullUrl = postUrl.Contains("http") ? postUrl : _getLegacyBaseAddress() + "/" + postUrl;
            return UploadFile(fullUrl, form, filePath, _getLegacyCookies());
        }

        //Used to upload several files or fake form data with multipart-form-data object
        public string Legacy_UploadFiles(string postUrl, Dictionary<string, string> form, List<FormUpload.FileParameter> files)
        {
            string fullUrl = postUrl.Contains("http") ? postUrl : _getLegacyBaseAddress() + "/" + postUrl;
            return UploadFiles(fullUrl, form, files, _getLegacyCookies());
        }

        public string Legacy_OptionsRequest(string postUrl)
        {
            string fullUrl = postUrl.Contains("http") ? postUrl : _getLegacyBaseAddress() + "/" + postUrl;
            return OptionsRequest(fullUrl, _getLegacyCookies());
        }

        private string _getLegacyBaseAddress()
        {
            var arr = Data.Get("config:legacy_url").Split('/');
            string baseUrl = string.Join("/", arr.Take(arr.Length - 1));
            return baseUrl;
        }

        // Returns cookies for legacy requests. If cookies = null - authorization will be performed first.
        private CookieContainer _getLegacyCookies()
        {
            return LEGACY_SITE_COOKIES ?? (LEGACY_SITE_COOKIES = _authLegacySite());
        }

        // Method which performs WS authorization and returns cookies for legacy site
        private CookieContainer _authLegacySite()
        {
            if (Data.Get("config:legacy_url").Equals(Data.Get("config:url")))
            {
                var browserCookies = GetCookies();
                if (browserCookies.Count > 0)
                {
                    return browserCookies;
                }
            }
            var url = Data.Get("config:legacy_url");
            string responseString;
            var baseAddress = new Uri(url);
            CookieContainer cookies = new CookieContainer();
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0");
                client.DefaultRequestHeaders.ConnectionClose = false;
                client.DefaultRequestHeaders.Host = url.Split('/')[2];
                var result = client.GetAsync(url).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            var formData = GetFormDataFromHTMLResponse(responseString);
            formData.Add("btnLoginV3$btn.x", "62");
            formData.Add("btnLoginV3$btn.y", "15");
            formData["txtUserName"] = Data.Get("config:username");
            formData["txtPassword"] = Data.Get("config:password");
            List<KeyValuePair<string, string>> form = formData.Select(x => new KeyValuePair<string, string>(x.Key, x.Value)).ToList();

            var encodedItems = form.Select(i => $"{WebUtility.UrlEncode(i.Key)}={WebUtility.UrlEncode(i.Value)}");
            var encodedContent = new StringContent(string.Join("&", encodedItems), null, "application/x-www-form-urlencoded");

            var req = new HttpRequestMessage(HttpMethod.Post, url) {Content = encodedContent};
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            using (var handler = new HttpClientHandler {CookieContainer = cookies})
            using (var client = new HttpClient(handler) {BaseAddress = baseAddress})
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
                client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.5");
                client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0");
                HttpResponseMessage result = client.SendAsync(req).Result;
                result.EnsureSuccessStatusCode();
                var response = result.Content.ReadAsStringAsync();
                responseString = response.Result;
            }
            return cookies;
        }

        public string GetBaseUrl()
        {
            string url = driver.Url;
            string[] urlArr = url.Split('/');
            string baseUrl = string.Concat(urlArr[0], "//", urlArr[2], "/", urlArr[3], "/");
            return baseUrl;
        }

        //Used for parse HTML response to fill dictionary with form data used for legacy site WS requests
        public static Dictionary<string, string> GetFormDataFromHTMLResponse(string html, bool removeButtons = true)
        {
            Dictionary<string, string> formData = new Dictionary<string, string>();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);

            //Fill inputs
            var inputs = doc.DocumentNode.SelectNodes("//form//input");
            if (inputs != null)
            {
                foreach (var input in inputs)
                {
                    var name = input.GetAttributeValue("name", null);
                    if (name == null) continue;
                    if (formData.ContainsKey(name)) continue;
                    if (removeButtons && name.Contains("$btn")) continue; // Skip adding buttons
                    formData.Add(input.GetAttributeValue("name", ""), input.GetAttributeValue("value", ""));
                }
            }

            //Fill selects
            var selections = doc.DocumentNode.SelectNodes("//form//select");
            if (selections != null)
            {
                foreach (var select in selections)
                {
                    string name = select.GetAttributeValue("name", "");
                    if (string.IsNullOrEmpty(name)) continue;
                    if (formData.ContainsKey(name)) continue;
                    var options = doc.DocumentNode.SelectNodes($"//select[@name='{name}']/option");
                    if (options == null) continue;
                    foreach (var option in options)
                    {
                        if (option.GetAttributeValue("selected", "").Equals("selected"))
                        {
                            formData.Add(name, option.GetAttributeValue("value", ""));
                            break;
                        }
                    }
                }
            }
            return formData;
        }

        //Get value of select or table list option
        public static string GetSelectOrTableValueId(string html, string value, string selectId)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            var select = doc.DocumentNode.SelectNodes($"//select[starts-with(@id,'{selectId}')]//option[contains(.,'{value}')]");
            if (select != null)
            {
                return select.First().GetAttributeValue("value", "");
            }
            var tableCell = doc.DocumentNode.SelectNodes($"//table[starts-with(@id,'{selectId}')]/tr[td/label[text()='{value}']]//input");
            if (tableCell != null)
            {
                return tableCell.First().GetAttributeValue("value", "");
            }
            return null;
        }

        //Get all option values in a dictionary
        public static Dictionary<string, string> GetSelectOptionValues(string html, string selectId)
        {
            Dictionary<string, string> ls = new Dictionary<string, string>();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            var options = doc.DocumentNode.SelectNodes($"//select[starts-with(@id,'{selectId}')]//option");
            foreach(var option in options)
            {
                ls.Add(option.InnerText, option.GetAttributeValue("value",String.Empty));
            }
            return ls;
        }

        public static string GetContentTypeByFileExtension(string file)
        {
            string result = "";
            switch (file.ToLower().Split('.').Last())
            {
                case "png":
                    result = "image/png";
                    break;
                case "jpg":
                    result = "image/jpeg";
                    break;
                case "doc":
                case "dot":
                    result = "application/msword";
                    break;
                case "docx":
                    result = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    break;
                case "dotx":
                    result = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
                    break;
                case "docm":
                case "dotm":
                    result = "application/vnd.ms-word.document.macroEnabled.12";
                    break;
                case "xls":
                case "xlt":
                case "xla":
                    result = "application/vnd.ms-excel";
                    break;
                case "xlsx":
                    result = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    break;
                case "xltx":
                    result = "application/vnd.openxmlformats-officedocument.spreadsheetml.template";
                    break;
                case "ppt":
                case "pot":
                case "pps":
                case "ppa":
                    result = "application/vnd.openxmlformats-officedocument.spreadsheetml.template";
                    break;
                case "pptx":
                    result = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                    break;
                case "potx":
                    result = "application/vnd.openxmlformats-officedocument.presentationml.template";
                    break;
                case "ppsx":
                    result = "application/vnd.openxmlformats-officedocument.presentationml.slideshow";
                    break;
            }
            return result;
        }
    }
}