using Appium.Pages.Settings;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.PageObjects;

namespace Appium.Pages
{
    public class FooterPage
    {

        public AndroidDriver<AndroidElement> _androidDriver;
        public FooterPage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        } 

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/navigation_settings")]
        private IWebElement _settingsMenu { get; set; }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/navigation_loads")]
        private IWebElement _loadsMenu { get; set; }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/navigation_history")]
        private IWebElement _historyMenu { get; set; }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/navigation_chat")]
        private IWebElement _chatMenu { get; set; }


        public LoadsPage NavigateToLoads()
        {
            _loadsMenu.Click();
            return new LoadsPage(_androidDriver);
        }

        public ChatPage NavigateToChat()
        {
            _chatMenu.Click();
            return new ChatPage(_androidDriver);
        }

        public SettingsPage NavigateToSettings()
        {
            _settingsMenu.Click();
            return new SettingsPage(_androidDriver);
        }



    }
}
