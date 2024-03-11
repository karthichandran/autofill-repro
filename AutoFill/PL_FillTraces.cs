using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Anticaptcha_example.Api;
using Microsoft.Playwright;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.DevTools.V85.Page;
using UglyToad.PdfPig.Fonts.TrueType.Names;

namespace AutoFill
{
    public class PL_FillTraces
    {
        private IPage page;

        public PL_FillTraces()
        {

        }
        private async Task Setup()
        {
            var playwright = await Playwright.CreateAsync();
            var browser =
                await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = false, SlowMo = 400, Timeout = 90000, Channel = "chrome" });
            page = await browser.NewPageAsync();
            await page.WaitForTimeoutAsync(1000);
        }

        public async Task<string> AutoFillForm16B(TdsRemittanceDto tdsRemittanceDto)
        {
            try
            {
                await FillLogin( tdsRemittanceDto);
                var reqNo =await RquestForm16B( tdsRemittanceDto);
               await page.CloseAsync();
                return reqNo;
            }
            catch (Exception e)
            {
                await page.CloseAsync();
                // MessageBox.Show("Request form16B Failed");
            }
            return "";
        }

        public async Task<string> AutoFillDownload(TdsRemittanceDto tdsRemittanceDto, string requestNo, DateTime dateOfBirth)
        {
            try
            {
             
                await FillLogin( tdsRemittanceDto);
                var fileName =await DownloadForm( requestNo, tdsRemittanceDto.CustomerPAN);
                if (fileName != "")
                {

                    UnzipFile unzipFile = new UnzipFile();
                    var filePath = unzipFile.extractFile(fileName, dateOfBirth.ToString("ddMMyyyy"));
                   await page.CloseAsync();
                    return filePath;

                }
                //else
                //    MessageBox.Show("Form is not yet generated");
            }
            catch (Exception e)
            {
                await page.CloseAsync();
                // MessageBox.Show("Download form Failed");
            }
            return null;
        }

        private async Task FillLogin( TdsRemittanceDto tdsRemittanceDto)
        {
            await Setup();
            await page.GotoAsync("https://www.tdscpc.gov.in/app/login.xhtml", new PageGotoOptions
            {
                WaitUntil = WaitUntilState.DOMContentLoaded
            });

            var logintype = page.Locator("#tpay");
           await logintype.ClickAsync();

           await page.WaitForTimeoutAsync(1000);
         

            var userId = page.Locator("#userId");
            //userId.SendKeys("ADMPC7474M");
            await userId.FillAsync(tdsRemittanceDto.CustomerPAN);
           
            var pwd = page.Locator("#psw");
            // pwd.SendKeys("Rana&123");
            pwd.FillAsync(tdsRemittanceDto.TracesPassword);

           //var pan = page.Locator("#tanpan");
           //// pwd.SendKeys("Rana&123");
           //await pan.FillAsync(tdsRemittanceDto.CustomerPAN);

            var captcha = await ReadCaptcha("captchaImg");
            if (captcha == "")
            {
                MessageBoxResult result = MessageBox.Show("Please fill the captcha and press OK button.", "Confirmation",
                    MessageBoxButton.OK, MessageBoxImage.Asterisk,
                    MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                var captchaInput = page.Locator("#captcha");
                await captchaInput.FillAsync(captcha);
            }

            var loginElm= page.Locator("#clickLogin");
           await loginElm.ClickAsync();

           await page.WaitForTimeoutAsync(2000);

            var confirmationChk = page.Locator("#Details");
           await confirmationChk.ClickAsync();
           await page.WaitForTimeoutAsync(2000);

            var confirmationBtn = page.Locator("#btn");
           await confirmationBtn.ClickAsync();
           await page.WaitForTimeoutAsync(2000);
        }

        private async Task<string> RquestForm16B(TdsRemittanceDto tdsRemittanceDto)
        {
            await page.GotoAsync("https://www.tdscpc.gov.in/app/tap/download16b.xhtml", new PageGotoOptions
            {
                WaitUntil = WaitUntilState.DOMContentLoaded
            });
           

            var formType = page.Locator("#formTyp");
            await formType.SelectOptionAsync(new []{ "26QB"});

            var assessmentYear = page.Locator("#assmntYear");
            await assessmentYear.SelectOptionAsync(new[] { tdsRemittanceDto.AssessmentYear });
           

            var actkNo = page.Locator("#ackNo");
           await actkNo.FillAsync(tdsRemittanceDto.ChallanAckNo);

            var panOfSeller = page.Locator("#panOfSeller");
            //panOfSeller.SendKeys("AJLPG4797J");
           await panOfSeller.FillAsync(tdsRemittanceDto.SellerPAN);

            var process = page.Locator("#clickGo");
          await process.ClickAsync();
          await page.WaitForTimeoutAsync(2000);

            var submitReq = page.Locator("#clickGo");
          await  submitReq.ClickAsync();
          await page.WaitForTimeoutAsync(2000);

            var requestTxt =await page.Locator("#hidReqId").GetAttributeAsync("value");
            return requestTxt;
        }

        private async Task<string> DownloadForm(string requestNo, string pan)
        {
            await page.GotoAsync("https://www.tdscpc.gov.in/app/tap/tpfiledwnld.xhtml", new PageGotoOptions
            {
                WaitUntil = WaitUntilState.DOMContentLoaded
            });
           

            var searchOpt = page.Locator("#search1");
           await searchOpt.ClickAsync();

            var requestTxt = page.Locator("#reqNo");
           await requestTxt.FillAsync(requestNo);

            var viewRequestBtn = page.Locator("#getListByReqId");
           await viewRequestBtn.ClickAsync();
           await page.WaitForTimeoutAsync(2000);
            var rows = page.Locator(".jqgrow");
            if (await rows.CountAsync() == 0)
                return "";

            var statusCell = rows.First.Locator("td").Nth(6);
            if ((await statusCell.InnerTextAsync()).Trim() != "Available")
                return "";

           await statusCell.ClickAsync();
           await page.WaitForTimeoutAsync(2000);
            var assessCell =(await rows.First.Locator("td").Nth(2).InnerTextAsync()).Trim();
            var ackNoCell = (await rows.First.Locator("td").Nth(4).InnerTextAsync()).Trim();

            var fileName = pan.Substring(0, 3) + "xxxxx" + pan.Substring(8, 2) + "_" + assessCell + "_" + ackNoCell + "-" + 1;

         

            //=====
            var waitForDownloadTask = page.WaitForDownloadAsync();
            var httpDownload = page.Locator("#downloadhttp");
            await httpDownload.ClickAsync();
            var download = await waitForDownloadTask;
            var filePath = DownloadFilePath(fileName);
            await download.SaveAsAsync(filePath);
           


            return fileName;
        }

        private string DownloadFilePath(string fileName)
        {
            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            var filePath = @downloadPath + "\\" + fileName + ".zip";
            return filePath;
        }
        protected async Task<string> ReadCaptcha( string captchaId)
        {
            var base64 = "";
            for (var i = 0; i < 20; i++)
            {
                var base64string = await page.EvaluateAsync<string>(@"()=>{
    var c = document.createElement('canvas');
    var ctx = c.getContext('2d');
    var img = document.getElementById('" + captchaId + "'); c.height=img.naturalHeight;c.width=img.naturalWidth; ctx.drawImage(img, 0, 0,img.naturalWidth, img.naturalHeight);var base64String = c.toDataURL(); return base64String;}");


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
