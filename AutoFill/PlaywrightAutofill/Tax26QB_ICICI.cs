using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Playwright;
using Microsoft.Win32;
using OpenQA.Selenium.DevTools.V85.Page;
using UglyToad.PdfPig.Fonts.TrueType.Names;

namespace AutoFill.PlaywrightAutofill
{
   public class Tax26QB_ICICI
   {
       private IPage page;
       private BankAccountDetailsDto _bankLogin;
       private string TransactionLog;
       private service svc;
        public Tax26QB_ICICI()
        {
            svc = new service();
            var path = GetChromePath();
            var proc1 = new ProcessStartInfo();
            string anyCommand = " --remote-debugging-port=9223 --user-data-dir=\"C:\\Users\\Demo\\chromepythondebug\"";
            proc1.UseShellExecute = true;
            proc1.FileName = path;
            proc1.Verb = "runas";
            proc1.Arguments = anyCommand;
            proc1.WindowStyle = ProcessWindowStyle.Normal;
            Process.Start(proc1);

            //var exitCode = Microsoft.Playwright.Program.Main(new[] { "install", "firefox" });
            //if (exitCode != 0)
            //{
            //    Console.WriteLine("Failed to install browsers");
            //    Environment.Exit(exitCode);
            //}
        }
        private async Task Setup()
        {
            var playwright = await Playwright.CreateAsync();
            //var browser =
            //    await playwright.Firefox.LaunchAsync(new BrowserTypeLaunchOptions { Headless = false, SlowMo = 400, Timeout = 90000});
             var browser = await playwright.Chromium.ConnectOverCDPAsync("http://localhost:9223");
            //var browser =
            //    await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = false, SlowMo = 400, Timeout = 90000,Channel = "chrome"});
            page = await browser.NewPageAsync();
            //await page.RouteAsync("**/automation-validator.min.js", async route =>await route.AbortAsync());
            await page.WaitForTimeoutAsync(1000);
        }
        private string GetChromePath()
        {
            const string suffix = @"Google\Chrome\Application\chrome.exe";
            var prefixes = new List<string> { Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) };
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesx86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (programFilesx86 != programFiles)
            {
                prefixes.Add(programFiles);
            }
            else
            {
                var programFilesDirFromReg = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion", "ProgramW6432Dir", null) as string;
                if (programFilesDirFromReg != null) prefixes.Add(programFilesDirFromReg);
            }

            prefixes.Add(programFilesx86);
            var path = prefixes.Distinct().Select(prefix => Path.Combine(prefix, suffix)).FirstOrDefault(File.Exists);
            return path;
        }

        public async Task<string> AutoFillForm26QB(AutoFillDto autoFillDto, string tds, string interest, string lateFee, BankAccountDetailsDto bankLogin, string transID,bool isResident)
        {
            try
            {
                TransactionLog = "";
                _bankLogin = bankLogin;
                await Setup();
                await LoginToIncomeTaxPortal(autoFillDto.eportal);
                var crn= await ProcessEportal(autoFillDto.eportal, isResident);
                await ProcessToBank( transID,autoFillDto.tab1.PanOfPayer);
              // var downloadStatus = await Download_DA(autoFillDto.tab1.PanOfPayer);
               // await LogOut();
                await page.CloseAsync();
                svc.SaveTransLog(int.Parse(transID), "Completed");
                return crn;
            }
            catch (Exception e) {
              // await LogOut();
               await page.CloseAsync();
                Console.WriteLine(e);
                svc.SaveTransLog(int.Parse(transID), TransactionLog);
                MessageBox.Show("Processing Form26QB Failed");
                return "";
            }
        }

        public async Task<string> AutoFillForm26QB_NoMsg(AutoFillDto autoFillDto, string tds, string interest, string lateFee, BankAccountDetailsDto bankLogin, string transID, bool isResident)
        {
            try
            {
                TransactionLog = "";
                _bankLogin = bankLogin;
                await Setup();
                await LoginToIncomeTaxPortal(autoFillDto.eportal);
                var crn = await ProcessEportal(autoFillDto.eportal, isResident);
                await ProcessToBank(transID, autoFillDto.tab1.PanOfPayer);
                //var downloadStatus = await Download_DA(autoFillDto.tab1.PanOfPayer);
                //await LogOut();
                await page.CloseAsync();
                svc.SaveTransLog(int.Parse(transID), "Completed");
                return crn;
            }
            catch (Exception e)
            {
                await page.CloseAsync();
                Console.WriteLine(e);
                svc.SaveTransLog(int.Parse(transID), TransactionLog);
                return "";
            }
        }

        public async Task<bool> DownloadChallanFromTaxPortal(AutoFillDto autoFillDto, int transID)
        {
            try
            {
                var svc = new service();
                var daObj = svc.GetDebitAdviceByClienttransId(transID);
                if (daObj == null || daObj.DebitAdviceID == 0)
                    return false;

                await Setup();
                await LoginToIncomeTaxPortal(autoFillDto.eportal);

                var isDownloaded =await DownloadChallan( daObj);
                await LogOut();
                await page.CloseAsync();
                return isDownloaded;
            }
            catch (Exception e)
            {
              // await  LogOut();
               await page.CloseAsync();
                return false;
                // throw;
            }
        }

        private async Task LoginToIncomeTaxPortal(Eportal eportal)
        {
            await page.GotoAsync("https://eportal.incometax.gov.in/iec/foservices/#/login", new PageGotoOptions
            {
                WaitUntil = WaitUntilState.DOMContentLoaded
            });

            TransactionLog = "Login failed";
            var userElm=await  page.WaitForSelectorAsync("#panAdhaarUserId");
           await userElm.FillAsync(eportal.LogInPan);

            var continueBtn = page.Locator(".large-button-primary");
           await continueBtn.ClickAsync();

           var confirmChk =await page.WaitForSelectorAsync(".mat-checkbox-layout",new PageWaitForSelectorOptions(){Timeout = 90000});
           await confirmChk.ClickAsync();
           await page.WaitForTimeoutAsync(500);

            var pwdElm = page.Locator("#loginPasswordField");
            await pwdElm.FillAsync(eportal.IncomeTaxPwd);
            await page.WaitForTimeoutAsync(500);

            continueBtn = page.Locator(".large-button-primary");
            await continueBtn.ScrollIntoViewIfNeededAsync();
            await page.WaitForTimeoutAsync(500);
           await continueBtn.ClickAsync();
            await page.WaitForTimeoutAsync(1000);
            var errorauth = page.Locator("mat-error");
            if (await errorauth.IsVisibleAsync())
            {
                await page.WaitForTimeoutAsync(2000);
                continueBtn = page.Locator(".large-button-primary");
                await continueBtn.ClickAsync();
            }

            await page.WaitForTimeoutAsync(2000);
          
            var loginHereBtn =  page.Locator(".primaryBtnMargin");
           if (await loginHereBtn.CountAsync()>0)
           {
               await loginHereBtn.ClickAsync();
               await page.WaitForTimeoutAsync(2000);
            }

           var itrBackBtn = await page.WaitForSelectorAsync(".previousIcon", new PageWaitForSelectorOptions() { Timeout = 60000 });
           //var itrBackBtn = page.Locator(".previousIcon");
           if (itrBackBtn != null && await itrBackBtn.IsEnabledAsync())
               await itrBackBtn.ClickAsync();

            TransactionLog = "Exit after Login";
            await page.WaitForURLAsync("**/dashboard",new PageWaitForURLOptions(){Timeout = 90000});
        }

        private async Task LogOut()
        {
            var userProfileBtn = page.Locator(".profileMenubtn");
            await userProfileBtn.ClickAsync();
            var receiptElm = page.Locator(".mat-menu-item").Nth(2);
            await receiptElm.ClickAsync();
        }

        private async Task<bool> DownloadChallan(DebitAdviceDto dto)
        {
            
            await page.GotoAsync(
                "https://eportal.incometax.gov.in/iec/foservices/#/dashboard/e-pay-tax/e-pay-tax-dashboard",new PageGotoOptions
                {
                    WaitUntil = WaitUntilState.DOMContentLoaded
                });
            await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
            var securityRiskBtn =await page.WaitForSelectorAsync("xpath=//*[@id='securityReasonPopup']/div/div/div[3]/button[2]",new PageWaitForSelectorOptions(){Timeout = 60000});
            await securityRiskBtn.ClickAsync();

            var tab = page.Locator("#mat-tab-label-0-2");
            await tab.ClickAsync();
            await page.EvaluateAsync("document.getElementById('ymPluginDivContainerInitial').remove();");
            //todo : remove popup

            var filterBtn = page.Locator(".filterMobile");
            await filterBtn.ClickAsync();

            var startDate = page.Locator(".mat-datepicker-toggle-default-icon").Nth(2);
            await startDate.ClickAsync();

            var datePart = new DatePart();
            datePart.Day = dto.PaymentDate.Value.Day;
            datePart.Month = dto.PaymentDate.Value.ToString("MMM");
            datePart.Year = dto.PaymentDate.Value.Year;

           await pickdate( datePart);

           var endDate = page.Locator(".mat-datepicker-toggle-default-icon").Nth(3);
           await endDate.ClickAsync();
           await pickdate( datePart);

           var filterSubmit = page.Locator(".primaryButton").Nth(5);
           await filterSubmit.ClickAsync();

            //
            //.mat-select-panel> mat-option>span

            var noOfRowElm = page.Locator("#mat-select-7");
            await noOfRowElm.ClickAsync();
            var optList = page.Locator(".mat-select-panel> mat-option>span");
            if (await optList.CountAsync()>0)
            {
                await optList.Nth(2).ClickAsync();
            }

            var gridContain = page.Locator(".ag-center-cols-container");
         
           var rows = gridContain.First.Locator(".ag-row");

           var isDownloaded = false;
           var rowCnt = await rows.CountAsync();
           for (var i=0;i<rowCnt;i++)
           {
               var row =  rows.Nth(i);
               var cells = row.Locator(".ag-cell");
               var cinNo =await cells.First.TextContentAsync();
                if (dto.CinNo == cinNo)
                {
                    var actionBtn = cells.Nth(6).Locator(".mat-icon-button");
                    await actionBtn.ClickAsync();
                    var waitForDownloadTask = page.WaitForDownloadAsync();

                    var receiptElm = page.Locator(".mat-menu-item").First;
                    await receiptElm.ClickAsync();
                    var download = await waitForDownloadTask;
                    var filePath = await DownloadFilePath( cinNo+ "_ChallanReceipt");
                    await download.SaveAsAsync(filePath);
                    isDownloaded = true;
                    break;
                }
            }
           return isDownloaded;
        }

        private async Task<string> DownloadFilePath(string fileName)
        {
            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            var filePath = @downloadPath + "\\" + fileName + ".pdf";
            return filePath;
        }

        private async Task<string> ProcessEportal( Eportal eportal, bool isResident)
        {
            await page.GotoAsync(
                "https://eportal.incometax.gov.in/iec/foservices/#/dashboard/e-pay-tax/26qb", new PageGotoOptions
                {
                    WaitUntil = WaitUntilState.DOMContentLoaded
                });

            await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
            var securityRiskBtn = await page.WaitForSelectorAsync("xpath=//*[@id='securityReasonPopup']/div/div/div[3]/button[2]", new PageWaitForSelectorOptions() { Timeout = 60000 });
            await securityRiskBtn.ClickAsync();

            TransactionLog = "Failed at residential selection";
            var residentStatus = page.Locator("xpath=//*[@id='mat-radio-5']/label");
            await residentStatus.ClickAsync();

            TransactionLog = "Failed at One / more buyer selection";
            if (!eportal.IsCoOwners)
            {
                var oneBuyer = page.Locator("xpath=//*[@id='mat-radio-9']/label");
               await oneBuyer.ClickAsync();
            }
            else
            {
                var moreBuyer = page.Locator("xpath=//*[@id='mat-radio-8']/label");
               await moreBuyer.ClickAsync();
            }

            //todo : scroll bottom
            TransactionLog = "Failed after filling buyers details";
            var continueBtn = page.Locator(".nextIcon").First;
            await continueBtn.ClickAsync();

            TransactionLog = "Failed at PAN of seller";
            var panSeller = page.Locator("xpath=//input[@formcontrolname='pan']").Nth(1);
           // var panSeller = page.Locator("#mat-input-8");
            await panSeller.FillAsync(eportal.SellerPan);
            await page.WaitForTimeoutAsync(1000);

            //var panSellerConfirm = page.Locator("#mat-input-31");
            TransactionLog = "Failed at confirm PAN of seller";
            var panSellerConfirm = page.Locator("xpath=//input[@formcontrolname='confirmPan']");
            await panSellerConfirm.FillAsync(eportal.SellerPan);
            await page.WaitForTimeoutAsync(1000);

            TransactionLog = "Failed at flat address of seller tab";
            //var flat = page.Locator("#mat-input-10");
            var flat = page.Locator("xpath=//input[@formcontrolname='flatAddress']").Nth(1);
            await flat.FillAsync(eportal.SellerFlat);
            await page.WaitForTimeoutAsync(500);

            TransactionLog = "Failed at road address  of seller tab";
            //var road = page.Locator("#mat-input-11");
            var road = page.Locator("xpath=//input[@formcontrolname='streetAddress']").Nth(1);
            await road.FillAsync(eportal.SellerRoad);
            await page.WaitForTimeoutAsync(500);

            TransactionLog = "Failed at pin code  of seller tab";
            // var pinCode =await page.WaitForSelectorAsync("#mat-input-33");
            var inx = isResident ? 1 : 0;
            var pinCode =await page.WaitForSelectorAsync("xpath=//input[@formcontrolname='pincode']>> nth="+inx);
            await pinCode.FillAsync(eportal.SellerPinCode.Trim());
            await page.WaitForTimeoutAsync(1000);

            TransactionLog = "Failed at mobile number";
            var mobileNo = page.Locator("#phone").Nth(1);
            await mobileNo.FillAsync(eportal.SellerMobile);
            await page.WaitForTimeoutAsync(1000);

            TransactionLog = "Failed at email";
            //var email = page.Locator("#mat-input-13");
            var email = page.Locator("xpath=//input[@formcontrolname='emailId']").Nth(1);
            await email.FillAsync(eportal.SellerEmail);

            TransactionLog = "Failed at one / more seller option";
            //var oneSeller = page.Locator("xpath=//*[@id='mat-radio-11']/label");
            var oneSeller = page.Locator("xpath=//mat-radio-group[@formcontrolname='isMultiple']").Nth(1).Locator("label").Nth(1);
            await oneSeller.ClickAsync();

            //todo : scroll bottom
            TransactionLog = "Failed after filling seller details";
            continueBtn = page.Locator(".large-button-primary").Nth(2);
            await continueBtn.ClickAsync();
            await page.WaitForTimeoutAsync(1000);

            TransactionLog = "Failed at property type selection";
            if (eportal.IsLand)
            {
                //var typeLand = page.Locator("xpath=//*[@id='mat-radio-17']/label");
                var typeLand = page.Locator("xpath=//mat-radio-group[@formcontrolname='propertyType']").Nth(0).Locator("label").Nth(0);
                await typeLand.ClickAsync();
            }
            else
            {
               // var typeBuild = page.Locator("xpath=//*[@id='mat-radio-18']/label");
                var typeBuild = page.Locator("xpath=//mat-radio-group[@formcontrolname='propertyType']").Nth(0).Locator("label").Nth(1);
                await typeBuild.ClickAsync();
            }

            TransactionLog = "Failed at property flat";
            //var propFlat = page.Locator("#mat-input-16");
            var propFlat = page.Locator("xpath=//input[@formcontrolname='flatAddress']").Nth(2);
            await propFlat.FillAsync(eportal.PropFlat);

            TransactionLog = "Failed at property road";
            //var propRoad = page.Locator("#mat-input-17");
            var propRoad = page.Locator("xpath=//input[@formcontrolname='streetAddress']").Nth(2);
            await propRoad.FillAsync(eportal.PropRoad);

            TransactionLog = "Failed at property pin";
            //var propPin = page.Locator("#mat-input-18");
            inx = isResident ? 2 : 1;
            var propPin = await page.WaitForSelectorAsync("xpath=//input[@formcontrolname='pincode']>> nth="+inx);
            await propPin.FillAsync(eportal.PropPinCode);

            TransactionLog = "Failed at date of agreement";
            var dateOfAgreement = page.Locator(".mat-datepicker-toggle-default-icon").First;
            await dateOfAgreement.ClickAsync();

            await pickdate(eportal.DateOfAgreement);

            TransactionLog = "Failed at property value";
            // var totalVal = page.Locator("#mat-input-22");
            var totalVal = page.Locator("xpath=//input[@formcontrolname='propertyValue']").Nth(0);
            await totalVal.FillAsync(eportal.TotalAmount.ToString());

            TransactionLog = "Failed at date of payment";
            var dateOfPay = page.Locator(".mat-datepicker-toggle-default-icon").Nth(1);
            await dateOfPay.ClickAsync();

            await pickdate(eportal.RevisedDateOfPayment);

            TransactionLog = "Failed at  payment type";
            if (eportal.paymentType == 1) // 1 is lumpsum
            {
                //var payTypeLump = page.Locator("xpath=//*[@id='mat-radio-21']/label");
                var payTypeLump = page.Locator("xpath=//mat-radio-group[@formcontrolname='paymentType']").Nth(0).Locator("label").Nth(1);
                await payTypeLump.ClickAsync();

                var isStamptDutyHiggerYes = page.Locator("xpath=//*[@id='mat-radio-26']/label");
              await  isStamptDutyHiggerYes.ClickAsync();
                //or
                var isStamptDutyHiggerNo = page.Locator("xpath=//*[@id='mat-radio-27']/label");
               await isStamptDutyHiggerNo.ClickAsync();
            }
            else
            {
                //var payTypeInstallment = page.Locator("xpath=//*[@id='mat-radio-20']/label");
                var payTypeInstallment = page.Locator("xpath=//mat-radio-group[@formcontrolname='paymentType']").Nth(0).Locator("label").Nth(0);
                await payTypeInstallment.ClickAsync();

                //var lastInstallmentNo = page.Locator("xpath=//*[@id='mat-radio-24']/label");
                var lastInstallmentNo = page.Locator("xpath=//mat-radio-group[@formcontrolname='lastInstallment']").Nth(0).Locator("label").Nth(1);
                await lastInstallmentNo.ClickAsync();

            }

            TransactionLog = "Failed at previous installment";
            var totalAmtPaidPreviously = page.Locator("xpath=//input[@formcontrolname='prevInstallment']");
            if(await totalAmtPaidPreviously.IsEnabledAsync())
            await totalAmtPaidPreviously.FillAsync(Math.Round(eportal.TotalAmountPaid).ToString());

            var amtPaidCurr = page.Locator("xpath=//input[@formcontrolname='amtPaidCurrently']");
            if (await amtPaidCurr.IsEnabledAsync())
                await amtPaidCurr.FillAsync(eportal.AmountPaid.ToString());

            var stampVal = page.Locator("xpath=//input[@formcontrolname='stampDutyValue']");
            if (await stampVal.IsEnabledAsync())
                await stampVal.FillAsync(eportal.StampDuty.ToString());

            //var tdsAmt = page.Locator("#mat-input-26");
            //var tdsAmt = page.Locator("xpath=//input[@formcontrolname='totalAmtPaid']");
            //if (await tdsAmt.IsEnabledAsync())
            //    await tdsAmt.FillAsync(eportal.Tds.ToString());

            TransactionLog = "Failed at TDS amount";
            var tdsAmt = page.Locator("xpath=//input[@formcontrolname='tdsAmnt']");
            if (await tdsAmt.IsEnabledAsync())
                await tdsAmt.FillAsync(Convert.ToInt32(eportal.Tds).ToString());

            // var interest = page.Locator("#mat-input-28");
            var interest = page.Locator("xpath=//input[@formcontrolname='interest']");
            if (await interest.IsEnabledAsync())
                await interest.FillAsync(Convert.ToInt32(eportal.Interest).ToString());

            //var fee = page.Locator("#mat-input-29");
            var fee = page.Locator("xpath=//input[@formcontrolname='others']");
            if (await fee.IsEnabledAsync())
                await fee.FillAsync(Convert.ToInt32(eportal.Fee).ToString());

            continueBtn = page.Locator(".large-button-primary").Nth(4);
            await continueBtn.ClickAsync();

            TransactionLog = "Failed at Bank name selection";
            var iciciNet = page.Locator("xpath=//img[@src='https://static.incometax.gov.in/iec/foservices/assets/iciciBank.png']");
            await iciciNet.ClickAsync();

            continueBtn = page.Locator(".large-button-primary").Nth(0);
            await continueBtn.ClickAsync();

            var payNowElm = page.Locator("button:has-text('Pay Now')").First;
            await payNowElm.ClickAsync();
            //todo : scroll to bottom
            var terms = page.Locator(".mat-checkbox-layout");
            await terms.ClickAsync();

           var submitToBank = page.Locator("xpath=//*[@id='SubmitToBank']/div/div/div[3]/button");
            await submitToBank.ClickAsync();

            TransactionLog = "Failed at Corporate User";
            var crnRow = page.Locator("table >tbody > tr > td > table > tbody > tr ");
             var crn = await crnRow.Nth(3).Locator("td").Nth(1).InnerTextAsync();
            var corporateUser = page.Locator("#CIB_11X_PROCEED");
            await corporateUser.ClickAsync();
            return crn;
        }

        private async Task ProcessToBank(  string transId,string pan)
        {
            TransactionLog = "Failed at bank login";
            var userIdTxt = page.Locator("#login-step1-userid");
            await userIdTxt.FillAsync(_bankLogin.UserName);

            var pwdTxt = page.Locator(".login-pwd");
            await pwdTxt.FocusAsync();
            await pwdTxt.FillAsync(_bankLogin.UserPassword);

            var proceedBtn = page.Locator("#VALIDATE_CREDENTIALS1");
            await proceedBtn.ClickAsync();

            if (!string.IsNullOrEmpty(transId))
            {
                var feeTxt = page.Locator(".type_RemarksMedium");
                await feeTxt.FillAsync(transId);
            }

            TransactionLog = "Failed at Grid";
            var gridAuth = page.Locator(".absmiddle").Nth(1);
            await gridAuth.ClickAsync();

            var continueBtn = page.Locator("#CONTINUE_PREVIEW");
            if (await continueBtn.CountAsync() > 0)
                await continueBtn.First.ClickAsync();

            await ProcessGridData();

            var submitBtn = page.Locator("#CONTINUE_SUMMARY");
            if (await submitBtn.CountAsync() > 0)
                await submitBtn.First.ClickAsync();

            await page.WaitForTimeoutAsync(1000);
        }

        private async Task<bool> Download_DA(string pan)
        {
            try
            {
                var waitForDownloadTask = page.WaitForDownloadAsync();
                var downloadBtn = page.Locator("#SINGLEPDF");
                var download = await waitForDownloadTask;
                var filePath = await DownloadFilePath(pan + "_");
                await download.SaveAsAsync(filePath);
                await downloadBtn.ClickAsync();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private async Task pickdate( DatePart date)
        {
            var pickerYear = page.Locator(".mat-calendar-period-button");  //
            await pickerYear.ClickAsync();
           
            var year = page.Locator("xpath=//div[@class='mat-calendar-body-cell-content' and contains(.,'" + date.Year + "')]");
            if (await year.CountAsync() == 0)
            {
                //mat-calendar-body-cell-content mat-calendar-body-today
                year = page.Locator("xpath=//div[@class='mat-calendar-body-cell-content mat-calendar-body-today' and contains(.,'" + date.Year + "')]");
            }
           await year.First.ClickAsync();
           
           var month = page.Locator("xpath=//div[@class='mat-calendar-body-cell-content' and contains(.,'" + date.Month.ToUpper() + "')]");
            if (await month.CountAsync() == 0)
            {
                //mat-calendar-body-cell-content mat-calendar-body-today
                month = page.Locator("xpath=//div[@class='mat-calendar-body-cell-content mat-calendar-body-today' and contains(.,'" + date.Month.ToUpper() + "')]");
            }
            await month.First.ClickAsync();
           
            var pickerDay = page.Locator("xpath=//div[@class='mat-calendar-body-cell-content' and contains(.,'" + date.Day + "')]");
            if (await pickerDay.CountAsync() == 0)
            {
                //mat-calendar-body-cell-content mat-calendar-body-today
                pickerDay = page.Locator("xpath=//div[@class='mat-calendar-body-cell-content mat-calendar-body-today' and contains(.,'" + date.Day + "')]");
            }
            await pickerDay.First.ClickAsync();
        }

        private async Task ProcessGridData()
        {
            Dictionary<string, string> grid = new Dictionary<string, string>();
            grid.Add("A", _bankLogin.LetterA.ToString());
            grid.Add("B", _bankLogin.LetterB.ToString());
            grid.Add("C", _bankLogin.LetterC.ToString());
            grid.Add("D", _bankLogin.LetterD.ToString());
            grid.Add("E", _bankLogin.LetterE.ToString());
            grid.Add("F", _bankLogin.LetterF.ToString());
            grid.Add("G", _bankLogin.LetterG.ToString());
            grid.Add("H", _bankLogin.LetterH.ToString());
            grid.Add("I", _bankLogin.LetterI.ToString());
            grid.Add("J", _bankLogin.LetterJ.ToString());
            grid.Add("K", _bankLogin.LetterK.ToString());
            grid.Add("L", _bankLogin.LetterL.ToString());
            grid.Add("M", _bankLogin.LetterM.ToString());
            grid.Add("N", _bankLogin.LetterN.ToString());
            grid.Add("O", _bankLogin.LetterO.ToString());
            grid.Add("P", _bankLogin.LetterP.ToString());

            var gridElms = page.Locator(".gridauth_input_cell_style");
            var firstLetter =await gridElms.Nth(0).InnerTextAsync();
            var secondLetter = await gridElms.Nth(1).InnerTextAsync();
            var thirdLetter =  await gridElms.Nth(2).InnerTextAsync();

            var firstInput = page.Locator(".gridauth_input_cell_style").Nth(3).Locator("input");
            await firstInput.FillAsync(grid[firstLetter]);
            var secondInput = page.Locator(".gridauth_input_cell_style ").Nth(4).Locator("input");
            await secondInput.FillAsync(grid[secondLetter]);
            var thirstInput = page.Locator(".gridauth_input_cell_style").Nth(5).Locator("input");
            await thirstInput.FillAsync(grid[thirdLetter]);

        }
    }
}
