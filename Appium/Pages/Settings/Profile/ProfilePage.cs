using Appium.Pages.Settings.Profile;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.PageObjects;

namespace Appium.Pages.Settings
{
    public class ProfilePage
    {
        public AndroidDriver<AndroidElement> _androidDriver;
        public ProfilePage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/tv_edit")]
        private IWebElement _profileEdit { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/tv_change_password")]
        private IWebElement _changePassword { get; set; }

        public EditProfilePage ClickOnEditProfile()
        {
            _profileEdit.Click();
            return new EditProfilePage(_androidDriver);
        }
    }
}
