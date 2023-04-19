using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace AutoFill
{
    public class FillTraces : Base
    {
        public static string AutoFillForm16B(TdsRemittanceDto tdsRemittanceDto) {
            try {
                var driver = GetChromeDriver();
                driver.Navigate().GoToUrl("https://www.tdscpc.gov.in/app/login.xhtml");
                WaitForReady(driver);
                FillLogin(driver, tdsRemittanceDto);
                var reqNo = RquestForm16B(driver, tdsRemittanceDto);
                   driver.Quit();
                return reqNo;
            }
            catch (Exception e) {
               // MessageBox.Show("Request form16B Failed");
            }
            return "";
        }

        public static string AutoFillDownload(TdsRemittanceDto tdsRemittanceDto,string requestNo,DateTime dateOfBirth )
        {
            try
            {
                var driver = GetChromeDriver();
                driver.Navigate().GoToUrl("https://www.tdscpc.gov.in/app/login.xhtml");
                WaitForReady(driver);
                FillLogin(driver, tdsRemittanceDto);
                var fileName= DownloadForm(driver, requestNo,tdsRemittanceDto.CustomerPAN);
                if (fileName != "")
                {

                    UnzipFile unzipFile = new UnzipFile();
                    var filePath=unzipFile.extractFile(fileName, dateOfBirth.ToString("ddMMyyyy"));
                    driver.Quit();
                    return filePath;

                }
                //else
                //    MessageBox.Show("Form is not yet generated");
            }
            catch (Exception e)
            {
               // MessageBox.Show("Download form Failed");
            }
            return null;
        }

        private static void FillLogin(IWebDriver webDriver, TdsRemittanceDto tdsRemittanceDto) {
            var logintype = webDriver.FindElement(By.Id("tpao"));
            logintype.Click();
            
            WaitForReady(webDriver);

            var userId = webDriver.FindElement(By.Id("userId"));
             //userId.SendKeys("ADMPC7474M");
            userId.SendKeys(tdsRemittanceDto.CustomerPAN);
            userId.SendKeys(Keys.Tab);
            var pwd = webDriver.FindElement(By.Id("psw"));
             // pwd.SendKeys("Rana&123");
            pwd.SendKeys(tdsRemittanceDto.TracesPassword);

           var captcha=  ReadCaptcha(webDriver, "captchaImg");
            if (captcha == "") {
                MessageBoxResult result = MessageBox.Show("Please fill the captcha and press OK button.", "Confirmation",
                                                         MessageBoxButton.OK, MessageBoxImage.Asterisk,
                                                         MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                var captchaInput = webDriver.FindElement(By.Id("captcha"));
                captchaInput.SendKeys(captcha);
            }

           

            WaitForReady(webDriver);
            webDriver.FindElement(By.Id("clickLogin")).Click();
            WaitForReady(webDriver);
            Thread.Sleep(1500);
            var confirmationChk= webDriver.FindElement(By.Id("Details"));
            confirmationChk.Click();
            WaitFor(webDriver, 2);
            var confirmationBtn = webDriver.FindElement(By.Id("btn"));
            confirmationBtn.Click();
            WaitForReady(webDriver);           
        }

        private static string RquestForm16B(IWebDriver webDriver, TdsRemittanceDto tdsRemittanceDto) {

            webDriver.Navigate().GoToUrl("https://www.tdscpc.gov.in/app/tap/download16b.xhtml");
            WaitForReady(webDriver);

            var formType = webDriver.FindElement(By.Id("formTyp"));
            var formTypeDDL = new SelectElement(formType);
            formTypeDDL.SelectByText("26QB");

            var assessmentYear = webDriver.FindElement(By.Id("assmntYear"));
            var assessmentYearDDL = new SelectElement(assessmentYear);
            // assessmentYearDDL.SelectByText("2020-21");
            assessmentYearDDL.SelectByText(tdsRemittanceDto.AssessmentYear);

            var actkNo = webDriver.FindElement(By.Id("ackNo"));
            actkNo.SendKeys(tdsRemittanceDto.ChallanAckNo);

            var panOfSeller = webDriver.FindElement(By.Id("panOfSeller"));
            //panOfSeller.SendKeys("AJLPG4797J");
            panOfSeller.SendKeys(tdsRemittanceDto.SellerPAN);

            var process = webDriver.FindElement(By.Id("clickGo"));
            process.Click();
            WaitForReady(webDriver);
            var submitReq = webDriver.FindElement(By.Id("clickGo"));
            submitReq.Click();
            WaitForReady(webDriver);

            var requestTxt = webDriver.FindElement(By.Id("hidReqId")).GetAttribute("value");
            return requestTxt;
        }

        private static string DownloadForm(IWebDriver webDriver, string requestNo,string pan)
        {
            webDriver.Navigate().GoToUrl("https://www.tdscpc.gov.in/app/tap/tpfiledwnld.xhtml");
            WaitForReady(webDriver);

            var searchOpt = webDriver.FindElement(By.Id("search1"));
            searchOpt.Click();
           
            var requestTxt = webDriver.FindElement(By.Id("reqNo"));
            requestTxt.SendKeys(requestNo);

            var viewRequestBtn = webDriver.FindElement(By.Id("getListByReqId"));
            viewRequestBtn.Click();
            WaitFor(webDriver, 2);
            var rows = webDriver.FindElements(By.ClassName("jqgrow"));
            if (rows.Count == 0)
                return "";
             
            var statusCell= rows[0].FindElements(By.TagName("td"))[6];
            if (statusCell.Text.Trim() != "Available")
                return "";
            statusCell.Click();
            WaitFor(webDriver, 1);
            var assessCell = rows[0].FindElements(By.TagName("td"))[2].Text.Trim();
            var ackNoCell = rows[0].FindElements(By.TagName("td"))[4].Text.Trim();

            var fileName = pan.Substring(0, 3) + "xxxxx" + pan.Substring(8, 2)+"_"+assessCell+"_"+ackNoCell+"-"+1;

            var httpDownload = webDriver.FindElement(By.Id("downloadhttp"));
            httpDownload.Click();
                       
            return fileName;
        }
    }
}
