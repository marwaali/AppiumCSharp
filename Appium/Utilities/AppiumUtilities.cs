using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Service;
using System;

namespace Appium.Utilities
{
    public class AppiumUtilities: Base
    {
        private AppiumLocalService _appiumLocalService;
        private readonly AndroidDriver<AndroidElement> _androidDriver;

        public AppiumUtilities(AppiumLocalService appiumLocalService, AndroidDriver<AndroidElement> androidDriver)
        {
            _appiumLocalService = appiumLocalService;
            _androidDriver = androidDriver;
        }
        public AppiumLocalService StartAppiumLocalService()
        {
            _appiumLocalService = new AppiumServiceBuilder().UsingAnyFreePort().Build();
            if (!_appiumLocalService.IsRunning)
            {
                _appiumLocalService.Start();
            }
            return _appiumLocalService;
        }
        public AppiumLocalService StartAppiumLocalService(int portNumber)
        {
            _appiumLocalService = new AppiumServiceBuilder().UsingPort(portNumber).Build();
            if (!_appiumLocalService.IsRunning)
            {
                _appiumLocalService.Start();
            }
            return _appiumLocalService;
        }
        public void StopAppiumLocalService()
        {
            _androidDriver.CloseApp();
            _appiumLocalService.Dispose();
        }
        public AndroidDriver<AndroidElement> InitializeAndroidNativeApp()
        {
            AppiumOptions option = new AppiumOptions();
            option.AddAdditionalCapability(MobileCapabilityType.DeviceName, "Android");
            option.AddAdditionalCapability(MobileCapabilityType.AutomationName, AutomationName.AndroidUIAutomator2);
            option.AddAdditionalCapability(AndroidMobileCapabilityType.AppPackage, "com.advansys.driver");
            option.AddAdditionalCapability(MobileCapabilityType.App, @"M:\APK\Driver.apk");
            //option.AddAdditionalCapability(AndroidMobileCapabilityType.AppPackage, "com.advansys.lms");
            //option.AddAdditionalCapability(AndroidMobileCapabilityType.AppActivity, "com.advansys.lms.mvp.vp.splash.SplashActivity");
            //option.AddAdditionalCapability(MobileCapabilityType.App, @"M:\APK\SED.apk");
            option.AddAdditionalCapability(MobileCapabilityType.PlatformName, "Android");

            AndroidDriver<AndroidElement> androidDriver = new AndroidDriver<AndroidElement>(_appiumLocalService, option, TimeSpan.FromSeconds(1000));
            return androidDriver;
        }
    }
}
