using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Appium.Pages
{
    public class LoadsPage
    {
        public AndroidDriver<AndroidElement> _androidDriver;
        public LoadsPage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/navigation_loads")]
        private IWebElement _loadsMenu { get; set; }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/toolbar_title")]
        private IWebElement _pageTitle { get; set; }

        public bool IsLoadsPageDisplayed()
        {
            return _loadsMenu.Displayed;
        }

        public string GetPageName()
        {
            return _pageTitle.Text;
        }

        public void ClickOnSkip()
        {
            Thread.Sleep(5000);
            (new TouchAction(_androidDriver)).Tap(570, 2241).Perform();
        }
    }
}
