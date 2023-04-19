using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using NUnit.Framework;
using NUnit.Framework.Interfaces;

namespace UIAutomation.Utils_Misc
{
    public class TestExecution
    {
        public static List<string> INPROGRESS = new List<string>();
        public static List<string> FINISHED = new List<string>();
        public static List<string> PASSED = new List<string>();
        public static List<string> SKIPPED = new List<string>();
        public static List<string> FAILED = new List<string>();
        
        static object locker = new object(); 

        public static void addTestResult(string testName, ResultState result)
        {
            if (result == ResultState.Error || result == ResultState.Failure)
            {
                addFailed(testName);
            }
            else if (result == ResultState.Ignored || result == ResultState.Skipped)
            {
                addSkipped(testName);
            }
            else if (result == ResultState.Success)
            {
                addPassed(testName);
            }
            addFinished(testName);
            removeInprogress(testName);
        }

        public static void FailTest(string testName)
        {
            addFailed(testName);
            addFinished(testName);
            removeInprogress(testName);
        }

        public static void SkipTest(string testName)
        {
            addSkipped(testName);
            addFinished(testName);
            removeInprogress(testName);
        }

        private static void addFailed(string testName)
        {
            lock (locker)
            {
                FAILED.Add(testName);
            }
        }
        
        private static void addFinished(string testName)
        {
            lock (locker)
            {
                FINISHED.Add(testName);
            }
        }
        
        public static void addInprogress(string testName)
        {
            lock (locker)
            {
                INPROGRESS.Add(testName);
            }
        }
        
        private static void addSkipped(string testName)
        {
            lock (locker)
            {
                SKIPPED.Add(testName);
            }
        }
        
        private static void addPassed(string testName)
        {
            lock (locker)
            {
                PASSED.Add(testName);
            }
        }
  
        private static void removeInprogress(string testName)
        {
            lock (locker)
            {
                INPROGRESS.Remove(testName);
            }
        }
        
    }
    

    public class Environment
    {
        private static bool _environment_set = false;

        public static string RootPath;
        public static string Env;
        public static string Browser;
        public static string ExecutionConfigFile;
        public static string TestDataFile;
        public static Dictionary<string, string> envParams;
        static object locker = new object();

        public static void SetEnvironment(TestParameters testParameters)
        {
            if (_environment_set) return;
            lock (locker)
            {
                if (_environment_set) return;
                RootPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)?.Substring(6).Split(new[] {"bin"}, StringSplitOptions.None)[0];
                Env = testParameters.Get("environment", ConfigurationManager.AppSettings.Get("Environment"));
                Browser = testParameters.Get("browser", ConfigurationManager.AppSettings.Get("Browser"));
                ExecutionConfigFile = testParameters.Get("ExecutionConfig", ConfigurationManager.AppSettings.Get("ExecutionConfig"));
                TestDataFile = testParameters.Get("TestData", new FileInfo(ConfigurationManager.AppSettings.Get("TestData")).Exists ? ConfigurationManager.AppSettings.Get("TestData") : RootPath + ConfigurationManager.AppSettings.Get("TestData"));
                envParams = ReadEnvConfig();
                _environment_set = true;
            }
        }

        private static Dictionary<string, string> ReadEnvConfig()
        {
            var configXml = XDocument.Load(RootPath + ExecutionConfigFile);
            var config = configXml.Descendants(Env);
            if (!config.Any())
            {
                throw new Exception("Environment [" + Env + "] not declared in the config");
            }

            Dictionary<string, string> configData = new Dictionary<string, string>();

            //Fill urls
            foreach (var url in config.Descendants("url"))
            {
                configData.Add($"config_{url.Attribute("ui").Value.ToLower()}:url", url.Value);
            }
            //Fill users
            foreach (var login in config.Descendants("LoginInfo"))
            {
                configData.Add($"config_{login.Attribute("persona").Value.ToLower()}:username", login.Attribute("username").Value);
                configData.Add($"config_{login.Attribute("persona").Value.ToLower()}:password", login.Attribute("password").Value);
            }
            //Fill ExchangeInfo
            foreach (var exchange in config.Descendants("EmailExchangeServerInformation"))
            {
                //configData.Add("config_mail:username", exchange.Attribute("username").Value);
                configData.Add("config_mail:username", exchange.Attribute("emailusername").Value);
                configData.Add("config_mail:password", exchange.Attribute("encryptedPassword").Value);
            }
            //Fill APIToken
            foreach (var apiKey in config.Descendants("API"))
            {
                configData.Add($"config:apikey", apiKey.Attribute("key")?.Value);
                configData.Add($"config:apiurl", apiKey.Attribute("url")?.Value);
            }

            return configData;
        }
    }
}