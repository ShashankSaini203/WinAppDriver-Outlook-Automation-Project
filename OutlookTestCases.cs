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

       //public static WindowsDriver<WindowsElement> OutlookSession;
        [ClassInitialize]
        public static void Initialize(TestContext context)
        {
            OutlookSessionSetup();
        }

        [ClassCleanup]
        public static void Dispose()
        {
            TearDown();
            WinApp.Close();
        }

        [TestMethod]
        public void OutlookNewEmail()
        {
            OutlookSessionCreate();
            outlookSession.FindElementByName("Home").Click();
            outlookSession.FindElementByName("New Email").Click();
            Thread.Sleep(TimeSpan.FromSeconds(3));
            try
            {
                OutlookSessionCreate();
                WindowsElement tbEmailTo = outlookSession.FindElementByXPath("//Edit[@ClassName='RichEdit20WPT'][@Name='To']");
                tbEmailTo.SendKeys("shashanksaini203@gmail.com");
                WindowsElement tbEmailCC = outlookSession.FindElementByXPath("//Edit[@ClassName='RichEdit20WPT'][@Name='Cc']");
                tbEmailCC.SendKeys("shashanksaini203@gmail.com");
                WindowsElement tbEmailSub = outlookSession.FindElementByXPath("//Edit[@ClassName='RichEdit20WPT'][@Name='Subject']");
                tbEmailSub.SendKeys("Test Automation using WinAppDriver");
                WindowsElement tareaEmailBody = outlookSession.FindElementByXPath("//*[@Name='Page 1 content'][@AutomationId='Body']");
                Actions action = new Actions(outlookSession);
                action.MoveToElement(tareaEmailBody,5,5).Click().Perform();
                action.MoveToElement(tareaEmailBody, 5, 5).SendKeys("Test Email from shashanksaini203@gmail.com").Perform();
                Thread.Sleep(TimeSpan.FromSeconds(1));
                
                WindowsElement btnEmailSend = outlookSession.FindElementByXPath("//Button[@ClassName ='Button'][@Name='Send']");
                btnEmailSend.Click();
            }
            catch(WebDriverException)
            {

            }

           
        }

    }
}
