using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;

namespace Appium.Pages.Settings.Profile
{
    public class EditProfilePage
    {
        public AndroidDriver<AndroidElement> _androidDriver;
        public EditProfilePage(AndroidDriver<AndroidElement> androidDriver)
        {
            PageFactory.InitElements(androidDriver, this);
            _androidDriver = androidDriver;
        }

        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/ed_first_name")]
        private IWebElement _editFirstName { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/ed_second_name")]
        private IWebElement _editsecondName { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/sp_country_code")]
        private IWebElement _editCountryCode { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/ed_phone")]
        private IWebElement _editPhoneNumber { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/ed_email")]
        private IWebElement _editEmail { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/sp_language")]
        private IWebElement _editLanguage { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/sp_locale_country")]
        private IWebElement _editLocaleCountry { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/sp_time_zone")]
        private IWebElement _editTimeZone { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/sp_license")]
        private IWebElement _editLicence { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/tv_save")]
        private IWebElement _save { get; set; }
        [FindsBy(How = How.Id, Using = "com.advansys.driver:id/tv_cancel")]
        private IWebElement _cancel { get; set; }

        public ProfilePage SaveEdit()
        {
            _save.Click();
            return new ProfilePage(_androidDriver);
        }

        public ProfilePage CancelEdit()
        {
            _cancel.Click();
            return new ProfilePage(_androidDriver);
        }

        private void SelectElementByValue(IWebElement element, string value)
        {
            SelectElement oSelect = new SelectElement(element);
            oSelect.SelectByValue(value);
        }

        public void UpdateProfileData(string firstName, string secondName, string phoneNumber, 
            string country, string email, string language, string localeCountry, string timeZone, string license)
        {
            _editFirstName.SendKeys(firstName);
            _editsecondName.SendKeys(secondName);
            _editPhoneNumber.SendKeys(phoneNumber);
            _editEmail.SendKeys(email);
            SelectElementByValue(_editCountryCode, country);
            SelectElementByValue(_editLanguage, language);
            SelectElementByValue(_editLocaleCountry, localeCountry);
            SelectElementByValue(_editTimeZone, timeZone);
            SelectElementByValue(_editLicence, license); 
        }
    }
}
