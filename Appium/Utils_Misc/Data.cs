using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using NUnit.Framework;
using NUnit.Framework.Internal;

namespace UIAutomation.Utils_Misc
{
    class Data
    {
        public static ConcurrentDictionary<string, Dictionary<string, string>> data { get; set; }
        
        // Next 'debugTimeStamp' variable is for overriding timestamp used in test scripts. If variable value is not empty timestamp
        // will be the value with underscore '_<value>', if empty - default current date/time will be used.
        private static string debugTimeStamp = "";
        public static string dateTimeStamp = debugTimeStamp != "" ? "_" + debugTimeStamp : DateTime.Now.ToString("_MMdd_HHmm");

        public static string Get(string key, string defaultValue = null)
        {
            var testData = (Dictionary<string, string>) TestExecutionContext.CurrentContext.CurrentTest.Properties.Get("TestData");
            if (testData.ContainsKey(key.ToLower()))
            {
                return getConfigDataReplacedValue(testData[key.ToLower()]);
            }
            if (defaultValue != null)
            {
                return defaultValue;
            }
            throw new InvalidDataException($"Test data doesn't contain key [{key}]");
        }

        public static void Set(string key, string value)
        {
            ((Dictionary<string, string>) TestExecutionContext.CurrentContext.CurrentTest.Properties.Get("TestData"))[key.ToLower()] = value;
        }

        //Used for getting (or creating if not created before) WebService object 
        public static WebService GetWebService()
        {
            if (!TestExecutionContext.CurrentContext.CurrentTest.Properties.ContainsKey("WebServiceObject"))
            {
                TestExecutionContext.CurrentContext.CurrentTest.Properties.Set("WebServiceObject", new WebService());
            }
            return (WebService) TestExecutionContext.CurrentContext.CurrentTest.Properties.Get("WebServiceObject");
        }

        public static string GetAppendDateTime(string key, string defaultValue = null)
        {
            return Get(key, defaultValue) + dateTimeStamp;
        }
        
        // Retrieve the post data from WebservicesPostString.xml For now it is assumed that this will be copied over into the build Debug folder from the Solution everytime it is run
        public static string GetPostData(string scriptname)
        {
            string applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); // This is explictly done so that when run from Jenkins the path is resolved correctly
            XDocument xmlDoc = XDocument.Load(Path.Combine(applicationPath, "WebServicesPostStrings.xml"));
            return xmlDoc.Root.Element(scriptname).Value;
        }
        
        // Replaces post data with the placeholder for prameters from a supplied data dictionary.
        // User has to carefull to supply all parameters in the disctionary supplied
        public static string ReplacePostDataByParam(string postData, Dictionary<string, string> param)
        {
            foreach (KeyValuePair<string, string> i in param)
            {
                postData = postData.Replace("<<" + i.Key + ">>", i.Value);
            }
            return postData;
        }

        #region supportingMethods
        // This method checks if there are any config values in the string and replaces them with their corresponding values
        private static string getConfigDataReplacedValue(string value)
        {
            // initialize regular expression
            string pattern = @"config_{0,1}\w*?:\w*";
            Regex regex = new Regex(pattern);

            if (value.Contains("config"))
            {
                MatchCollection results = regex.Matches(value);
                foreach (Match item in results)
                {
                    if (data.ContainsKey(item.Value))
                    {
                        value = value.Replace(item.Value, data[TestContext.CurrentContext.Test.ID][item.Value]);
                    }
                }
            }
            return value;
        }
        #endregion supportingMethods
    }
}
