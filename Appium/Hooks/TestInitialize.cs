using Appium.Utilities;
using NUnit.Framework;
using System;

namespace Appium.Hooks
{
    [TestFixture]
    public class TestInitialize : Base
    {
        [SetUp]
        public void InitializeTest()
        {
            AndroidContext = StartAppiumServerForNative();
            AndroidContext.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        }

        [TearDown]
        public void CleanUp()
        {
            AppiumUtilities.StopAppiumLocalService();
        }
    }
}
