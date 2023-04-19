using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Appium.Pages
{
    public class LoginPage
    {
        public AndroidDriver<AndroidElement> _androidDriver;
        public LoginPage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/email")]
        private IWebElement _userEmail { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/password")]
        private IWebElement _userPassword { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/email_sign_in_button")]
        private IWebElement _loginButton { get; set; }
        [FindsBy(How = How.Id, Using = "com.android.permissioncontroller:id/permission_allow_foreground_only_button")]
        private IWebElement _allPermission { get; set; }
        [FindsBy(How = How.Id, Using = "com.android.permissioncontroller:id/permission_allow_button")]
        private IWebElement _allowPermission { get; set; }
        [FindsBy(How = How.TagName, Using = "tip_route_info")]
        private IWebElement _toolTip { get; set; }
        [FindsBy(How = How.CssSelector, Using = "com.advansys.driver:id/tv_next")]
        private IWebElement _toolTipNext { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/tv_error")]
        private IWebElement _errorText { get; set; }



        public LoadsPage loginToDriverApp(string userName, string password)
        {
            _userEmail.SendKeys(userName);
            _userPassword.SendKeys(password);
            _loginButton.Click();
            return new LoadsPage(_androidDriver);
        }

        public void AcceptPermissions()
        {
            for (int i = 0; i< 3; i++)
            {
                _allPermission.Click();
            }
            for (int i = 0; i< 2; i++)
            {
                _allowPermission.Click();
            }
        }

        public string GetErrorText()
        {
            return _errorText.Text;
        }

        public bool IsUserNameFieldDisplayed()
        {
            return _userEmail.Displayed;
        }
    }
}
