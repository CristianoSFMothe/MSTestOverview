using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace MSTestOverview
{
    [TestClass]
    public class AlarmsAndClockSmokeTests
    {

        static WindowsDriver<WindowsElement> sessionAlarms;
        private static TestContext objTestContext;

        [ClassInitialize]
        public static void PrepareForTestingAlarms(TestContext testContext)
        {
            Debug.WriteLine("Hello ClassInitialize");

            AppiumOptions capCalc = new AppiumOptions();

            capCalc.AddAdditionalCapability("app", "Microsoft.WindowsAlarms_8wekyb3d8bbwe!App");

            sessionAlarms = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), capCalc);

            objTestContext = testContext;

        }

        [ClassCleanup]
        public static void CleanupAfterAllAlarmsTests()
        {
            Debug.WriteLine("Hello ClassCleanup");

            if (sessionAlarms != null)
            {
                sessionAlarms.Quit();
            }
        }

        [TestInitialize]
        public void BeforeATest()
        {
            Debug.WriteLine("Before a test, calling TestInitialize");
        }

        [TestCleanup]
        public void AfterATest()
        {
            Debug.WriteLine("After a test, calling TestCleanup");
        }

        [TestMethod]
        public void JustAnotherTest()
        {
            Debug.WriteLine("Hello another test.");
        }

        [TestMethod]
        public void TestAlarmsAndCLockIsLaunchingSuccessfully()
        {

            Debug.WriteLine("Hello TestAlarmsIsLaunchingSuccessfully!");

            Assert.AreEqual("Clock", sessionAlarms.Title, false,
                $"Actual title doesn't match expected title: {sessionAlarms.Title}");

        }

        [TestMethod, DataSource("System.Data.OleDb",
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Documentos\Workspace\Cursos\Udemy\WinAppDriver\clockTest.xlsx;Extended Properties=""Excel 8.0;HDR = YES"";",
            "Clocks$", DataAccessMethod.Sequential)]

        public void VerifyNewClockCabBeAdded()
        {
            // clique no botão Relógio mundial na seção do aplicativo
            sessionAlarms.FindElementByAccessibilityId("ClockDetailListView").Click();

            sessionAlarms.FindElementByName("Add a new city").Click();

            //System.Threading.Thread.Sleep(1000);

            WebDriverWait waitForMe = new WebDriverWait(sessionAlarms, TimeSpan.FromSeconds(10));

            var txtLocation = sessionAlarms.FindElementByName("Enter a location");

            waitForMe.Until(pred => txtLocation.Displayed);

            txtLocation.SendKeys("Lisbon, Portugal");

            txtLocation.SendKeys(Keys.Enter);

            sessionAlarms.FindElementByName("Add").Click();

            var listClock = sessionAlarms.FindElementByAccessibilityId("ClockDetailListView");

            var clockItems = listClock.FindElementsByTagName("ListItem");

           Debug.WriteLine($"Total tiles found: {clockItems.Count}");

            bool wasClockTiteFound = false;

            WindowsElement tileFound = null;

            foreach(WindowsElement clockTite in clockItems)
            {
                if (clockTite.Text.StartsWith("Lisbon, Portugal"))
                {
                    wasClockTiteFound = true;
                    Debug.WriteLine("Clock found.");
                    tileFound = clockTite;
                    break;
                }
                
            }

            Assert.IsTrue(wasClockTiteFound, "No clock tile found");

            Actions actionForRightClick = new Actions(sessionAlarms);

            actionForRightClick.MoveToElement(tileFound);

            actionForRightClick.Click();

            actionForRightClick.ContextClick();

            actionForRightClick.Perform();

            AppiumOptions capDesktop = new AppiumOptions();

            capDesktop.AddAdditionalCapability("app", "Root");

            WindowsDriver<WindowsElement> sessionDesktop = new WindowsDriver<WindowsElement>(
                new Uri("http://127.0.0.1:4723"), capDesktop
                );

            var contextItemDelete = sessionDesktop.FindElementByAccessibilityId("ContextMenuDelete");

            WebDriverWait desktopWaitForMe = new WebDriverWait(sessionAlarms, TimeSpan.FromSeconds(10));

            waitForMe.Until(pred => contextItemDelete.Displayed);

            contextItemDelete.Click();

        }

    }
}

