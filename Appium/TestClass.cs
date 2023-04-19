using Appium.Hooks;
using Appium.Pages;
using Appium.Pages.Settings;
using Appium.Pages.Settings.Profile;
using NUnit.Framework;

namespace Appium
{
    [TestFixture]
    public class TestClass: TestInitialize
    {
        [Test]
        public void LoginWithValidData()
        {
            LoginPage login = new LoginPage(AndroidContext);
            LoadsPage loads = login.loginToDriverApp("driver2@marwa.com", "Test!1234567");
            login.AcceptPermissions();
            loads.ClickOnSkip();
            Assert.IsTrue(loads.IsLoadsPageDisplayed());
            Assert.AreEqual("My Loads", loads.GetPageName());
        }

        [Test]
        public void LoginWithInValidData()
        {
            LoginPage login = new LoginPage(AndroidContext);
            login.loginToDriverApp("xxx@marwa.com", "Test!1234567");
            Assert.AreEqual("Invalid username or password", login.GetErrorText());
        }

        [Test]
        public void CheckForLogout()
        {
            LoginPage login = new LoginPage(AndroidContext);
            LoadsPage loads = login.loginToDriverApp("driver2@marwa.com", "Test!1234567");
            login.AcceptPermissions();
            loads.ClickOnSkip();
            FooterPage footer = new FooterPage(AndroidContext);
            SettingsPage settings= footer.NavigateToSettings();
            LoginPage login1 = settings.ClickOnLogout();
            Assert.IsTrue(login1.IsUserNameFieldDisplayed());
        }

        [Test]
        public void CheckForEditProfile()
        {
            LoginPage login = new LoginPage(AndroidContext);
            LoadsPage loads = login.loginToDriverApp("driver2@marwa.com", "Test!1234567");
            login.AcceptPermissions();
            loads.ClickOnSkip();
            FooterPage footer = new FooterPage(AndroidContext);
            SettingsPage settings = footer.NavigateToSettings();
            ProfilePage profile = settings.NavigateToProfilePage();
            EditProfilePage editProfile = profile.ClickOnEditProfile();
            //editProfile.UpdateProfileData();


        }
    }
}