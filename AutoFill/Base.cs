using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;
using Anticaptcha_example.Api;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using WebDriverManager.DriverConfigs.Impl;

namespace AutoFill
{
   public class Base
    {
        protected static void WaitForReady(IWebDriver webDriver)
        {
            TimeSpan timeSpan = TimeSpan.FromSeconds(120);
            WebDriverWait wait = new WebDriverWait(webDriver, timeSpan);
            wait.Until(driver => {
                bool isAjaxFinished = (bool)((IJavaScriptExecutor)driver).
                    ExecuteScript("return jQuery.active == 0");
                try
                {
                    var loader = driver.FindElement(By.ClassName("loader-mask")).GetAttribute("style");
                    Console.WriteLine(loader);
                    return loader.Split(':')[1] == " none;";
                }
                catch
                {
                    return isAjaxFinished;
                }
            });
        }
        protected static IWebElement GetElementById(IWebDriver webDriver, string id)
        {
            WaitFor(webDriver, 2);
            IWebElement element = new WebDriverWait(webDriver, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementIsVisible(By.Id(id)));
            return element;
        }
        protected static IWebElement GetElementByClass(IWebDriver webDriver, string cls)
        {
            WaitFor(webDriver, 2);
            IWebElement element = new WebDriverWait(webDriver, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementIsVisible(By.ClassName(cls)));
            return element;
        }
        protected static IWebElement GetElementByXpath(IWebDriver webDriver, string path)
        {
            WaitFor(webDriver, 2);
            IWebElement element = new WebDriverWait(webDriver, TimeSpan.FromSeconds(60)).Until(ExpectedConditions.ElementIsVisible(By.XPath(path)));
            return element;
        }

        protected static void ScrollToBottom(IWebDriver webDriver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)webDriver;
            js.ExecuteScript("window.scrollTo(0, 0)");
        }
        protected static void WaitFor(IWebDriver webDriver, int inSeconds = 0)
        {
            // webDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(inSeconds);
            Thread.Sleep(inSeconds * 1000);
        }

        protected static IWebDriver GetChromeDriver()
        {
            try
            {
                //Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe /T");

                //Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver");
                //foreach (var chromeDriverProcess in chromeDriverProcesses)
                //{
                //    chromeDriverProcess.Kill();
                //}

                ChromeOptions options = new ChromeOptions();
                options.AddArgument("--no-sandbox");
                options.AddArgument("--disable-infobars");
                options.AddArgument("--disable-dev-shm-usage");
                options.AddArgument("--start-maximized");

                // options.BinaryLocation = AppDomain.CurrentDomain.BaseDirectory+"chromedriver.exe";
                // options.AddArgument("--remote-debugging-port=9222");

                //var driver = new ChromeDriver(AppDomain.CurrentDomain.BaseDirectory, options);

               // new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
                ChromeDriver driver = new ChromeDriver( options);

               // ChromeDriver driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);

                return driver;
                //var ieDriver = GetIEDriver();
                //return ieDriver;
            }
            catch (Exception ex) {
                throw ex;
            }
        }

        protected static IWebDriver GetIEDriver() {

           // InternetExplorerDriver ie = new InternetExplorerDriver();
            EdgeDriver edge = new EdgeDriver();
            return edge;
        }

        protected static string ReadCaptcha(IWebDriver webDriver,string captchaId)
        {

            var jsExecuter = (IJavaScriptExecutor)webDriver;
            var base64 = "";
            for (var i = 0; i < 20; i++)
            {
                var base64string = jsExecuter.ExecuteScript(@"
    var c = document.createElement('canvas');
    var ctx = c.getContext('2d');
    var img = document.getElementById('" + captchaId + "'); c.height=img.naturalHeight;c.width=img.naturalWidth; ctx.drawImage(img, 0, 0,img.naturalWidth, img.naturalHeight);var base64String = c.toDataURL(); return base64String;") as string;

                base64 = base64string.Split(',').Last();

                if (string.IsNullOrEmpty(base64))
                {
                    Thread.Sleep(3000);
                }
                else
                    break;
            }


            var ClientKey = "f35d396e27db69a278ead2739cb85e99";
            var captcha = "";
            var api = new ImageToText
            {
                ClientKey = ClientKey,
                BodyBase64 = base64
            };

            if (!api.CreateTask())
            {
                MessageBox.Show(api.ErrorMessage, "Error");
            }
            else if (!api.WaitForResult())
            {
                MessageBox.Show("Could not solve the captcha.", "Error");
                //  DebugHelper.Out("Could not solve the captcha.", DebugHelper.Type.Error);
            }
            else
            {
                captcha = api.GetTaskSolution().Text;
                // DebugHelper.Out("Result: " + api.GetTaskSolution().Text, DebugHelper.Type.Success);
            }
            return captcha;
        }
    }


}
