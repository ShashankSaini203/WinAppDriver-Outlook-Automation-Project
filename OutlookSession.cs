using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAutomation
{
    public class OutlookSession
    {

        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string OutlookAppId = @"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE";
        static bool outlookWindow=false;
        public static Process WinApp;

        public static WindowsDriver<WindowsElement> outlookSession;
        public static WindowsDriver<WindowsElement> desktopSession;
        public static void OutlookSessionSetup()
        {
            WinApp = Process.Start(@"..\..\Windows Application Driver\WinAppDriver.exe");
            AppiumOptions desktopCapabilities = new AppiumOptions();
            desktopCapabilities.AddAdditionalCapability("app", "Root");
            desktopSession = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), desktopCapabilities);

            try
            {
                outlookWindow = desktopSession.FindElementByClassName("rctrl_renwnd32").Displayed;
            }
            catch(WebDriverException)
            {
                outlookWindow = false;
            }
            
            
            if (!outlookWindow)
            {
                
                AppiumOptions outlookCapabilities = new AppiumOptions();
                outlookCapabilities.AddAdditionalCapability("app", OutlookAppId);
                outlookSession = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), outlookCapabilities);
                outlookSession.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                Thread.Sleep(TimeSpan.FromSeconds(2));
            }
            else
            {

                OutlookSessionCreate();

            }



        }

        public static void OutlookSessionCreate()
        {
            WindowsElement outlookWindowElement = desktopSession.FindElementByClassName("rctrl_renwnd32");
            string OutlookTopLevelWindowHandle = OutlookTopLevelWindowHandle = (int.Parse(outlookWindowElement.GetAttribute("NativeWindowHandle"))).ToString("x");
            AppiumOptions outlookCapabilities = new AppiumOptions();
            outlookCapabilities.AddAdditionalCapability("appTopLevelWindow", OutlookTopLevelWindowHandle);
            outlookSession = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), outlookCapabilities);
            Thread.Sleep(TimeSpan.FromSeconds(2));
            
        }

        public static void TearDown()
        {
            if (outlookSession != null)
            {
                outlookSession.Quit();
                outlookSession = null;

            }
            if (desktopSession != null)
            {
                desktopSession.Quit();
                desktopSession = null;
            }
        }
    }
}
