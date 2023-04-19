using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.PageObjects;

namespace Appium.Pages
{
    public class ChatPage
    {
        public AndroidDriver<AndroidElement> _androidDriver;
        public ChatPage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        }
    }
}
