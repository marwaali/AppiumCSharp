using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.PageObjects;

namespace Appium.Pages.Settings
{
    public class SettingsPage
    {
        public AndroidDriver<AndroidElement> _androidDriver;
        public SettingsPage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/ll_logout")]
        private IWebElement _logout { get; set; }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/ll_profile")]
        private IWebElement _profile { get; set; }

        public LoginPage ClickOnLogout()
        {
            _logout.Click();
            return new LoginPage(_androidDriver);
        }

        public ProfilePage NavigateToProfilePage()
        {
            _profile.Click();
            return new ProfilePage(_androidDriver);
        }

    }
}
