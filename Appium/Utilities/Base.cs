using AventStack.ExtentReports;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Service;

namespace Appium.Utilities
{
    public class Base
    {
        public AppiumLocalService AppiumServiceContext;
        public AndroidDriver<AndroidElement> AndroidContext;
        public ExtentReports _extent;
        public  ExtentTest _test;
        public string projDir;


        public AppiumUtilities AppiumUtilities => new AppiumUtilities(AppiumServiceContext, AndroidContext);

        public AndroidDriver<AndroidElement> StartAppiumServerForNative()
        {
            AppiumServiceContext = AppiumUtilities.StartAppiumLocalService();
            AndroidContext = AppiumUtilities.InitializeAndroidNativeApp();
            return AndroidContext;
        }

        

    }
}
