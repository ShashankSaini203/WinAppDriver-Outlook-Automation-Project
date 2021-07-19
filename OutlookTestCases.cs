using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using System;
using System.Threading;

namespace OutlookAutomation
{

    //*****************************************************************************************************************************
    //
    // DEVELOPER - SHASHANK SAINI
    // DEVELOPER'S EMAIL- SHASHANKSAINI203@GMAIL.COM
    // GITHUB - https://github.com/ShashankSaini203
    // LinkedIn - https://in.linkedin.com/in/shashank-saini-203
    // This code is developed in an effort to illustrate the uses and benefits of WinAppDriver for Windows Automation. 
    // In this example, I have created an automation of Outlook application. MSTest testing framework is used in this code.
    // This code is for educational and learning purpose only.

    //*****************************************************************************************************************************

    [TestClass]
    public class OutlookTestCases : OutlookSession
    {

        [ClassInitialize]
        public static void Initialize(TestContext context)
        {
            OutlookSessionSetup();
        }

        [ClassCleanup]
        public static void Dispose()
        {
            TearDown();

        }

        [TestMethod]
        public void SendNewEmail()
        {
            OutlookSessionCreate();
            outlookSession.FindElementByName("Home").Click();
            outlookSession.FindElementByName("New Email").Click();
            Thread.Sleep(TimeSpan.FromSeconds(3));
            try
            {

                OutlookSessionCreate();
                WindowsElement tbEmailTo = outlookSession.FindElementByXPath("//*[@Name='Form Regions']/Edit[@ClassName='RichEdit20WPT'][@Name='To']");
                tbEmailTo.SendKeys("shashanksaini203@gmail.com");
                WindowsElement tbEmailCC = outlookSession.FindElementByXPath("//Edit[@ClassName='RichEdit20WPT'][@Name='Cc']");
                tbEmailCC.SendKeys("shashanksaini203@gmail.com");
                WindowsElement tbEmailSub = outlookSession.FindElementByXPath("//Edit[@ClassName='RichEdit20WPT'][@Name='Subject']");
                tbEmailSub.SendKeys("Test Automation");
                WindowsElement tareaEmailBody = outlookSession.FindElementByXPath("//*[@Name='Page 1 content'][@AutomationId='Body']");
                Actions action = new Actions(outlookSession);
                action.MoveToElement(tareaEmailBody, 5, 5).Click().Perform();
                action.MoveToElement(tareaEmailBody, 5, 5).SendKeys("This is a test Email").Perform();
                Thread.Sleep(TimeSpan.FromSeconds(1));

                WindowsElement btnEmailSend = outlookSession.FindElementByXPath("//Button[@ClassName ='Button'][@Name='Send']");
                btnEmailSend.Click();
            }
            catch (WebDriverException)
            {

            }


        }

        [TestMethod]
        public void SetupTeamsMeeting()
        {
            string meetingTitle = "MS Teams Meeting automation by WinAppDriver";
            string sendToRequired = "shashanksaini203@gmail.com";
            string sendToOptional = "shashanksaini203@gmail.com";
            string startDate = "1-Aug-2021";  //dd-mmm-yyyy
            string startTime = "03:00 PM";
            string endDate = "1-Aug-2021";
            string endTime = "03:30 PM";
            string meetingBody = "Hi All," + Keys.Enter + Keys.Enter + "Please join this meeting.";


            outlookSession.FindElementByName("Home").Click();
            Thread.Sleep(TimeSpan.FromSeconds(1));

            outlookSession.FindElementByName("New Items").Click();
            Thread.Sleep(TimeSpan.FromSeconds(1));
            outlookSession.FindElementByName("Teams Meeting").Click();
            Thread.Sleep(TimeSpan.FromSeconds(10));
            var allWindowHandles1 = outlookSession.WindowHandles;
            outlookSession.SwitchTo().Window(allWindowHandles1[0]);

            try
            {
                Actions action1 = new Actions(outlookSession);
                WindowsElement tbTeamsMeetingTitle = outlookSession.FindElementByName("Title");
                action1.MoveToElement(tbTeamsMeetingTitle, 200, 10).Click();
                action1.SendKeys(tbTeamsMeetingTitle, meetingTitle).Perform();

                Actions action2 = new Actions(outlookSession);
                WindowsElement tbTeamsMeetingRequired = outlookSession.FindElementByAccessibilityId("4109");
                action2.MoveToElement(tbTeamsMeetingRequired).Click();
                action2.SendKeys(tbTeamsMeetingRequired, sendToRequired);
                action2.SendKeys(tbTeamsMeetingRequired, Keys.Tab).Perform();


                Actions action3 = new Actions(outlookSession);
                WindowsElement tbTeamsMeetingOptional = outlookSession.FindElementByAccessibilityId("4110");
                action3.MoveToElement(tbTeamsMeetingOptional).Click();
                action3.SendKeys(tbTeamsMeetingOptional, sendToOptional);
                action3.SendKeys(tbTeamsMeetingOptional, Keys.Tab).Perform();


                WindowsElement txtStartDate = outlookSession.FindElementByName("Start date");
                txtStartDate.SendKeys(Keys.LeftControl + "a");
                txtStartDate.SendKeys(Keys.Backspace);
                txtStartDate.SendKeys(startDate);
                txtStartDate.SendKeys(Keys.Enter);
                Thread.Sleep(TimeSpan.FromSeconds(2));
                outlookSession.Keyboard.SendKeys(startTime);

                WindowsElement txtEndDate = outlookSession.FindElementByName("End date");
                txtEndDate.SendKeys(Keys.LeftControl + "a");
                txtEndDate.SendKeys(Keys.Backspace);
                txtEndDate.SendKeys(endDate);
                txtEndDate.SendKeys(Keys.Enter);
                Thread.Sleep(TimeSpan.FromSeconds(2));
                outlookSession.Keyboard.SendKeys(endTime);


                WindowsElement btnLocation = outlookSession.FindElementByName("Location");
                btnLocation.SendKeys(Keys.Tab);
                outlookSession.Keyboard.SendKeys(Keys.Tab);
                outlookSession.Keyboard.SendKeys(Keys.Tab);
                outlookSession.Keyboard.SendKeys(meetingBody);
                WindowsElement btnSend = outlookSession.FindElementByName("Send");
                btnSend.Click();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
