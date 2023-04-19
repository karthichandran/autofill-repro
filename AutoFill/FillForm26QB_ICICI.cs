using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace AutoFill
{
    public class FillForm26QB_ICICI : Base
    {
        static BankAccountDetailsDto _bankLogin;
        public static bool AutoFillForm26QB(AutoFillDto autoFillDto, string tds, string interest, string lateFee, BankAccountDetailsDto bankLogin, string transID)
        {
            try
            {

                _bankLogin = bankLogin;// rgan31
                                       // _bankLogin =new BankAccountDetailsDto{ UserName="reprosri",UserPassword="Repro&123"}; // Note : sri ram account
                                       // _bankLogin = new BankAccountDetailsDto { UserName = "579091011.RGANESH", UserPassword = "Rajalara@123" }; 
                                       // _bankLogin = new BankAccountDetailsDto { UserName = "579091011.VIJAYALA", UserPassword = "Sriram@123" }; 

                var driver = GetChromeDriver();
               
                driver.Navigate().GoToUrl("https://eportal.incometax.gov.in/iec/foservices/#/login");
                WaitForReady(driver);             

                ProcessEportal(driver , autoFillDto.eportal);
                ProcessToBank(driver, tds, interest, lateFee, transID);

                driver.Quit();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                MessageBox.Show("Processing Form26QB Failed");
                return false;
                // throw;
            }
        }

        //Note : both autofillform26q should be same functionality
        public static bool AutoFillForm26QB_NoMsg(AutoFillDto autoFillDto, string tds, string interest, string lateFee, BankAccountDetailsDto bankLogin, string transID)
        {
            try
            {
               

                _bankLogin = bankLogin;

                var driver = GetChromeDriver();

                driver.Navigate().GoToUrl("https://eportal.incometax.gov.in/iec/foservices/#/login");
                WaitForReady(driver);

                ProcessEportal(driver, autoFillDto.eportal);
                ProcessToBank(driver, tds, interest, lateFee, transID);

                driver.Quit();
                return true;
            }
            catch (Exception e)
            {
                return false;
                // throw;
            }
        }

        public static bool DownloadChallanFromTaxPortal(AutoFillDto autoFillDto,  int transID)
        {
            try
            {
                var svc = new service();
                var daObj = svc.GetDebitAdviceByClienttransId( transID);


                var driver = GetChromeDriver();

                driver.Navigate().GoToUrl("https://eportal.incometax.gov.in/iec/foservices/#/login");
                WaitForReady(driver);

                LoginToIncomeTaxPortal(driver, autoFillDto.eportal);
                DownloadChallan(driver, daObj);

                driver.Quit();
                return true;
            }
            catch (Exception e)
            {
                return false;
                // throw;
            }
        }

        private static void LoginToIncomeTaxPortal(IWebDriver webDriver, Eportal eportal) {
            //var userId = webDriver.FindElement(By.Id("panAdhaarUserId"));
            var userId = GetElementById(webDriver, "panAdhaarUserId");
            userId.SendKeys(eportal.LogInPan);
            // userId.SendKeys("AHUPB2786K");

            //var continueBtn = webDriver.FindElement(By.ClassName("large-button-primary"));
            var continueBtn = GetElementByClass(webDriver, "large-button-primary");
            if (!continueBtn.Displayed)
            {
                WaitFor(webDriver, 2);
                continueBtn = GetElementByClass(webDriver, "large-button-primary");
            }

            continueBtn.Click();
            WaitForReady(webDriver);

            // var confirmChk = webDriver.FindElement(By.ClassName("mat-checkbox-layout"));
            var confirmChk = GetElementByClass(webDriver, "mat-checkbox-layout");
            confirmChk.Click();

            // var pwdElm = webDriver.FindElement(By.Id("loginPasswordField"));
            var pwdElm = GetElementById(webDriver, "loginPasswordField");
            pwdElm.SendKeys(eportal.IncomeTaxPwd);
            //pwdElm.SendKeys("Rama1976$$");

            //continueBtn = webDriver.FindElement(By.ClassName("large-button-primary"));
            continueBtn = GetElementByClass(webDriver, "large-button-primary");
            continueBtn.Click();
            WaitFor(webDriver, 3);
            WaitForReady(webDriver);
        }

        private static void DownloadChallan(IWebDriver webDriver, DebitAdviceDto dto) {

            WaitForReady(webDriver);
            // Navigate to efile tab
            webDriver.Navigate().GoToUrl("https://eportal.incometax.gov.in/iec/foservices/#/dashboard/e-pay-tax/e-pay-tax-dashboard");
            WaitFor(webDriver, 2);
            var securityRiskBtn = GetElementByXpath(webDriver, "//*[@id='securityReasonPopup']/div/div/div[3]/button[2]");
            securityRiskBtn.Click();
            WaitForReady(webDriver);
            WaitFor(webDriver, 2);
            //payment tab
            var tab = GetElementById(webDriver, "mat-tab-label-0-2");
            tab.Click();                     

            //button[.filterMobile]
            var filterBtn = GetElementByClass(webDriver, "filterMobile");
            filterBtn.Click();
            WaitFor(webDriver, 2);

            var startDate = webDriver.FindElements(By.ClassName("mat-datepicker-toggle-default-icon"))[2];
            startDate.Click();

            var datePart = new DatePart();
            datePart.Day = dto.PaymentDate.Value.Day;
            datePart.Month = dto.PaymentDate.Value.ToString("MMM");
            datePart.Year = dto.PaymentDate.Value.Year;

            pickdate(webDriver, datePart);

            var endDate = webDriver.FindElements(By.ClassName("mat-datepicker-toggle-default-icon"))[3];
            endDate.Click();
            pickdate(webDriver, datePart);

           
            var filterSubmit = webDriver.FindElements(By.ClassName("primaryButton"))[4];
            filterSubmit.Click();

            var gridContain = webDriver.FindElement(By.ClassName("ag-center-cols-container"));
            var rows= gridContain.FindElements(By.ClassName("ag-row"));
          
            foreach (var row in rows) {
                var cells = row.FindElements(By.ClassName("ag-cell"));
                var cinNo = cells[0].Text;
                if (dto.CinNo == cinNo) { 
                var actionBtn=cells[6].FindElement(By.ClassName("mat-icon-button"));
                    actionBtn.Click();

                    var receiptElm = webDriver.FindElements(By.ClassName("mat-menu-item"))[0]; 
                    receiptElm.Click();
                }
            }
           

        }


        private static void ProcessEportal(IWebDriver webDriver, Eportal eportal) {
            //var userId = webDriver.FindElement(By.Id("panAdhaarUserId"));
            var userId = GetElementById(webDriver,"panAdhaarUserId");
            userId.SendKeys(eportal.LogInPan);
           // userId.SendKeys("AHUPB2786K");

            //var continueBtn = webDriver.FindElement(By.ClassName("large-button-primary"));
            var continueBtn = GetElementByClass(webDriver, "large-button-primary");
            if (!continueBtn.Displayed)
            {
                WaitFor(webDriver, 2);
                continueBtn = GetElementByClass(webDriver, "large-button-primary");
            }

            continueBtn.Click();
            WaitForReady(webDriver);

            // var confirmChk = webDriver.FindElement(By.ClassName("mat-checkbox-layout"));
            var confirmChk = GetElementByClass(webDriver, "mat-checkbox-layout");
            confirmChk.Click();

           // var pwdElm = webDriver.FindElement(By.Id("loginPasswordField"));
            var pwdElm = GetElementById(webDriver, "loginPasswordField");
            pwdElm.SendKeys(eportal.IncomeTaxPwd);
            //pwdElm.SendKeys("Rama1976$$");

            //continueBtn = webDriver.FindElement(By.ClassName("large-button-primary"));
            continueBtn = GetElementByClass(webDriver, "large-button-primary");
            continueBtn.Click();
            WaitFor(webDriver, 3);
            WaitForReady(webDriver);

           // webDriver.Navigate().GoToUrl("https://eportal.incometax.gov.in/iec/foservices/#/dashboard/e-pay-tax/e-pay-tax-dashboard");

            //large-button-secondary defualtButtonGap newPaymentbuttonRight ng-star-inserted

            //Close warning popup
            //var securityRiskBtn = webDriver.FindElement(By.XPath("//*[@id='securityReasonPopup']/div/div/div[3]/button[2]"));
            //securityRiskBtn.Click();
            //WaitForReady(webDriver);
            //WaitForReadyEportal(webDriver);

            //var newPayBtn = webDriver.FindElement(By.ClassName("large-button-secondary"));
            //newPayBtn.Click();

            // var procedd26qbBtn = webDriver.FindElements(By.ClassName("large-button-secondary"));
            // procedd26qbBtn[5].Click();

            //or 
            webDriver.Navigate().GoToUrl("https://eportal.incometax.gov.in/iec/foservices/#/dashboard/e-pay-tax/26qb");
           // var securityRiskBtn = webDriver.FindElement(By.XPath("//*[@id='securityReasonPopup']/div/div/div[3]/button[2]"));
            var securityRiskBtn = GetElementByXpath(webDriver, "//*[@id='securityReasonPopup']/div/div/div[3]/button[2]");
            securityRiskBtn.Click();
            WaitForReady(webDriver);
            WaitFor(webDriver, 2); 

           // var residentStatus = webDriver.FindElement(By.XPath("//*[@id='mat-radio-2']/label"));
            var residentStatus = GetElementByXpath(webDriver, "//*[@id='mat-radio-2']/label");
            residentStatus.Click();
            WaitFor(webDriver, 1);

            //*[@id="mat-radio-5"]/label
            if (!eportal.IsCoOwners)
            {
               // var oneBuyer = webDriver.FindElement(By.XPath("//*[@id='mat-radio-5']/label"));
                var oneBuyer = GetElementByXpath(webDriver, "//*[@id='mat-radio-5']/label");
                oneBuyer.Click();
            }
            else {
                //var moreBuyer = webDriver.FindElement(By.XPath("//*[@id='mat-radio-6']/label"));
                var moreBuyer = GetElementByXpath(webDriver, "//*[@id='mat-radio-6']/label");
                moreBuyer.Click();
            }

            ScrollToBottom(webDriver);
            //continueBtn = webDriver.FindElement(By.ClassName("large-button-primary"));
            continueBtn = GetElementByClass(webDriver, "large-button-primary");
            continueBtn.Click();
            WaitFor(webDriver, 3);
            //tab 2
            //var panSeller = webDriver.FindElement(By.Id("mat-input-9"));
            var panSeller = GetElementById(webDriver, "mat-input-9");
            panSeller.SendKeys(eportal.SellerPan);

           // var panSellerConfirm = webDriver.FindElement(By.Id("mat-input-33"));
            var panSellerConfirm = GetElementById(webDriver, "mat-input-34");
            panSellerConfirm.SendKeys(eportal.SellerPan);

           // var flat = webDriver.FindElement(By.Id("mat-input-11"));
            var flat = GetElementById(webDriver, "mat-input-11");
            flat.SendKeys(eportal.SellerFlat);

           // var road = webDriver.FindElement(By.Id("mat-input-12"));
            var road = GetElementById(webDriver, "mat-input-12");
            road.SendKeys(eportal.SellerRoad);

           // var pincode = webDriver.FindElement(By.Id("mat-input-34"));
            var pincode = GetElementById(webDriver, "mat-input-36");
            pincode.SendKeys(eportal.SellerPinCode.Trim());
            WaitFor(webDriver, 2);


            var mobileNo = webDriver.FindElements(By.Id("phone"))[1];
            mobileNo.SendKeys(eportal.SellerMobile);
            //mobileNo.SendKeys("9865321452");

            var email = GetElementById(webDriver, "mat-input-15");
            email.SendKeys(eportal.SellerEmail);
            //email.SendKeys("etes@fg.lo");


            var oneSeller = GetElementByXpath(webDriver, "//*[@id='mat-radio-11']/label");
            oneSeller.Click();
            //or
            //var moreSeller = webDriver.FindElement(By.XPath("//*[@id='mat-radio-12']/label"));
            //moreSeller.Click();

            ScrollToBottom(webDriver);
            continueBtn = webDriver.FindElements(By.ClassName("large-button-primary"))[2];
            continueBtn.Click();
            WaitFor(webDriver, 3);
            // tab 3
            if (eportal.IsLand)
            {
                var typeLand = GetElementByXpath(webDriver, "//*[@id='mat-radio-17']/label");
                typeLand.Click();
            }
            else
            {
                var typeBuild = GetElementByXpath(webDriver, "//*[@id='mat-radio-18']/label");
                typeBuild.Click();
            }

            var propFlat = GetElementById(webDriver, "mat-input-16");
            propFlat.SendKeys(eportal.PropFlat);

            //var propRoad = webDriver.FindElement(By.Id("mat-input-17"));
            var propRoad = GetElementById(webDriver, "mat-input-17");
            propRoad.SendKeys(eportal.PropRoad);


            //var propPin = webDriver.FindElement(By.Id("mat-input-18"));
            var propPin = GetElementById(webDriver, "mat-input-18");
            propPin.SendKeys(eportal.PropPinCode);
            WaitFor(webDriver, 2);
            //var propPostOffice = webDriver.FindElement(By.Id("mat-select-7"));
            //var propArea = webDriver.FindElement(By.Id("mat-select-8"));


            var dateOfAgreement  = webDriver.FindElements(By.ClassName("mat-datepicker-toggle-default-icon"))[0];
            dateOfAgreement.Click();
            pickdate(webDriver, eportal.DateOfAgreement);

              var totalVal = GetElementById(webDriver, "mat-input-22");
            totalVal.SendKeys(eportal.TotalAmount.ToString());

            var dateOfPay = webDriver.FindElements(By.ClassName("mat-datepicker-toggle-default-icon"))[1];
            dateOfPay.Click();
            pickdate(webDriver, eportal.RevisedDateOfPayment);
           
            // var dateOfDeduction = webDriver.FindElements(By.ClassName("mat-datepicker-toggle-default-icon"))[2];
            
            

            ////*[@id="mat-radio-20"]/label

            if (eportal.paymentType == 1) // 1 is lumpsum
            {
                var payTypeLump = GetElementByXpath(webDriver, "//*[@id='mat-radio-21']/label");
                payTypeLump.Click();

                var isStamptDutyHiggerYes = GetElementByXpath(webDriver, "//*[@id='mat-radio-26']/label");
                isStamptDutyHiggerYes.Click();
                //or
                var isStamptDutyHiggerNo = GetElementByXpath(webDriver, "//*[@id='mat-radio-27']/label");
                isStamptDutyHiggerNo.Click();
            }
            else {
                var payTypeInstallment = GetElementByXpath(webDriver, "//*[@id='mat-radio-20']/label");
                payTypeInstallment.Click();

                var lastInstallmentNo = GetElementByXpath(webDriver, "//*[@id='mat-radio-24']/label");
                lastInstallmentNo.Click();

            }

            var totalAmtPaidPreviously = GetElementById(webDriver, "mat-input-37");
            if(totalAmtPaidPreviously.Enabled)
                totalAmtPaidPreviously.SendKeys(Math.Round(eportal.TotalAmountPaid).ToString());
               // totalAmtPaidPreviously.SendKeys(eportal.TotalAmountPaid.ToString());

            var amtPaidCurr = GetElementById(webDriver, "mat-input-38");
            if (amtPaidCurr.Enabled)
                amtPaidCurr.SendKeys(eportal.AmountPaid.ToString());

            var stampVal = GetElementById(webDriver, "mat-input-39");
            if(stampVal.Enabled)
            stampVal.SendKeys(eportal.StampDuty.ToString());

            var tdsAmt = GetElementById(webDriver, "mat-input-26");
            if(tdsAmt.Enabled)
            tdsAmt.SendKeys(eportal.Tds.ToString());

            var interest = GetElementById(webDriver, "mat-input-28");
            if (interest.Enabled)
                interest.SendKeys(eportal.Interest.ToString());

            var fee = GetElementById(webDriver, "mat-input-29");
            if (fee.Enabled)
                fee.SendKeys(eportal.Fee.ToString());

            ScrollToBottom(webDriver);
            WaitFor(webDriver, 1);
            continueBtn = webDriver.FindElements(By.ClassName("large-button-primary"))[4];
            continueBtn.Click();
            WaitFor(webDriver, 3);
            //tab 3

            //*[@id="mat-radio-35"]/label

            //
           // var iciciNet = GetElementByXpath(webDriver, "//*[@id='mat-radio-41']/label");
            var iciciNet = GetElementByXpath(webDriver, "//img[@src='https://static.incometax.gov.in/iec/foservices/assets/iciciBank.png']");
            iciciNet.Click();
            WaitFor(webDriver, 1);

            continueBtn = webDriver.FindElements(By.ClassName("large-button-primary"))[0];
            continueBtn.Click();
            WaitFor(webDriver, 3);

            ScrollToBottom(webDriver);
            WaitFor(webDriver, 1);
            var payBtn = GetElementByXpath(webDriver, "//button[contains(.,'Pay Now')]");
            payBtn.Click();

            //mat-checkbox-layout
            var terms = GetElementByClass(webDriver, "mat-checkbox-layout");
            terms.Click();

            //*[@id="SubmitToBank"]/div/div/div[3]/button
            var submitToBank = GetElementByXpath(webDriver, "//*[@id='SubmitToBank']/div/div/div[3]/button");
            submitToBank.Click();
            WaitFor(webDriver, 1);

            var corporateUser = GetElementById(webDriver, "CIB_11X_PROCEED");
            corporateUser.Click();
            WaitFor(webDriver, 1);

            //var alert = webDriver.SwitchTo().Alert();
            //alert.Accept(); // or alert.dismiss()
        }

        ////*[@id="mat-datepicker-2"]/div/mat-multi-year-view/table/tbody/tr[2]/td[2]/div
        private static void pickdate(IWebDriver webDriver, DatePart date) {
            var pickerYear = webDriver.FindElement(By.ClassName("mat-calendar-period-button"));  //
            pickerYear.Click();
            WaitFor(webDriver, 2);
            var year = webDriver.FindElements(By.XPath("//div[@class='mat-calendar-body-cell-content' and contains(.,'" + date.Year + "')]"));
            if (year.Count() == 0) {
                //mat-calendar-body-cell-content mat-calendar-body-today
                year = webDriver.FindElements(By.XPath("//div[@class='mat-calendar-body-cell-content mat-calendar-body-today' and contains(.,'" + date.Year + "')]"));
            }
            year[0].Click();
            WaitFor(webDriver, 1);
            var month = webDriver.FindElements(By.XPath("//div[@class='mat-calendar-body-cell-content' and contains(.,'" + date.Month.ToUpper() + "')]"));
            if (month.Count() == 0)
            {
                //mat-calendar-body-cell-content mat-calendar-body-today
                month = webDriver.FindElements(By.XPath("//div[@class='mat-calendar-body-cell-content mat-calendar-body-today' and contains(.,'" + date.Month.ToUpper() + "')]"));
            }
            month[0].Click();
            WaitFor(webDriver, 1);
            //var pickerMonth = webDriver.FindElements(By.ClassName("mat-calendar-body-cell-content"))[0];
            //pickerMonth.Click();
            var pickerDay = webDriver.FindElements(By.XPath("//div[@class='mat-calendar-body-cell-content' and contains(.,'" + date.Day + "')]"));
            if ( pickerDay.Count()==0)
            {
                //mat-calendar-body-cell-content mat-calendar-body-today
                pickerDay = webDriver.FindElements(By.XPath("//div[@class='mat-calendar-body-cell-content mat-calendar-body-today' and contains(.,'" + date.Day + "')]"));
            }
            pickerDay[0].Click();
        }

      
        private static void ProcessToBank(IWebDriver webDriver, string tds, string interest, string lateFee, string transId)
        {
            if (_bankLogin == null)
            {
                var result = MessageBox.Show("Bank login details not available", "Confirmation",
                                                         MessageBoxButton.OK, MessageBoxImage.Asterisk,
                                                         MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }
            //WaitFor(webDriver, 3);

            //var payBtn = webDriver.FindElement(By.Id("CIB_11X_PROCEED"));
            //payBtn.Click();
            //WaitForReady(webDriver);
            WaitFor(webDriver, 3);
            var userIdTxt = webDriver.FindElement(By.Id("login-step1-userid"));
            userIdTxt.SendKeys(_bankLogin.UserName);
            var pwdTxt = webDriver.FindElement(By.Id("AuthenticationFG.ACCESS_CODE"));
            pwdTxt.SendKeys(_bankLogin.UserPassword);
            var proceedBtn = webDriver.FindElement(By.Id("VALIDATE_CREDENTIALS1"));
            proceedBtn.Click();
            WaitForReady(webDriver);
            //var incomeTaxTxt = webDriver.FindElement(By.Id("TranRequestManagerFG.TAX_AMOUNT_STR"));
            //incomeTaxTxt.SendKeys(tds);
            //WaitFor(webDriver, 1);
            //if (!string.IsNullOrEmpty(interest))
            //{
            //    var interestTxt = webDriver.FindElement(By.Id("TranRequestManagerFG.INTEREST_AMOUNT_STR"));
            //    interestTxt.SendKeys(interest);
            //}
            //WaitFor(webDriver, 1);
            //if (!string.IsNullOrEmpty(lateFee))
            //{
            //    var feeTxt = webDriver.FindElement(By.Id("TranRequestManagerFG.OTHER_FEE_AMT_STR"));
            //    feeTxt.SendKeys(lateFee);
            //}

            WaitFor(webDriver, 1);
            if (!string.IsNullOrEmpty(transId))
            {
                var feeTxt = webDriver.FindElement(By.Id("TranRequestManagerFG.PMT_RMKS"));
                feeTxt.SendKeys(transId);
            }

            var gridAuth = webDriver.FindElements(By.Id("TranRequestManagerFG.AUTH_MODES"));
            gridAuth[1].Click();

            var continueBtn = webDriver.FindElements(By.Id("CONTINUE_PREVIEW"));
            if (continueBtn.Count > 0)
            {
                continueBtn[0].Click();
                WaitForReady(webDriver);
            }
            ProcessGridData(webDriver);

            var submitBtn = webDriver.FindElements(By.Id("CONTINUE_SUMMARY"));
            if (submitBtn.Count > 0)
            {
                submitBtn[0].Click();
                WaitForReady(webDriver);
            }

            var downloadBtn = webDriver.FindElement(By.Id("SINGLEPDF"));
            downloadBtn.Click();
            WaitFor(webDriver, 3);
        }

        private static void ProcessToBank_hdfc(IWebDriver webDriver, string tds, string interest, string lateFee, string transId)
        {
            if (_bankLogin == null)
            {
                var result = MessageBox.Show("Bank login details not available", "Confirmation",
                                                         MessageBoxButton.OK, MessageBoxImage.Asterisk,
                                                         MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }
            WaitFor(webDriver, 3);

            var payBtn = webDriver.FindElements(By.XPath("//a[contains(.,'Click Here')]"))[0];
            payBtn.Click();
            //WaitForReady(webDriver);
            WaitFor(webDriver, 3);//
            var frame = webDriver.FindElements(By.Name("bottom_frame"))[0];
            webDriver.SwitchTo().Frame(frame);
            ////*[@id="pageBody"]/div[1]/form/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/input
            ////html/body/div[1]/form/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/input
            var userIdTxt = webDriver.FindElement(By.XPath("//html/body/div[1]/form/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/input"));
            // elms = webDriver.FindElements(By.ClassName("form-control"));
            //  var userIdTxt = webDriver.FindElement(By.ClassName("form-control"));
            userIdTxt.SendKeys(_bankLogin.UserName);

            var continueBtn = webDriver.FindElements(By.XPath("//a[contains(.,'CONTINUE')]"))[0];
            continueBtn.Click();
            // WaitForReady(webDriver);
            WaitFor(webDriver, 3);
            frame = webDriver.FindElements(By.Name("bottom_frame"))[0];
            webDriver.SwitchTo().Frame(frame);
            var pwdTxt = webDriver.FindElements(By.Id("fldPasswordDispId"));
            pwdTxt[0].SendKeys(_bankLogin.UserPassword);

            var loginBtn = webDriver.FindElements(By.XPath("//a[contains(.,'LOGIN')]"))[0];
            loginBtn.Click();
            WaitFor(webDriver, 3);

            var acctElm = webDriver.FindElements(By.Name("selAcct"));
            var acctDDl = new SelectElement(acctElm[0]);
            acctDDl.SelectByIndex(1);

            //basic tax
            var tdsInt = Math.Round(Convert.ToDecimal(tds), MidpointRounding.AwayFromZero);
            var incomeTaxTxt = webDriver.FindElement(By.Name("fldBasicTax"));
            incomeTaxTxt.Clear();
            incomeTaxTxt.SendKeys(tdsInt.ToString());
            WaitFor(webDriver, 1);

            if (!string.IsNullOrEmpty(interest))
            {
                var interestInt = Math.Round(Convert.ToDecimal(interest), MidpointRounding.AwayFromZero);
                var interestTxt = webDriver.FindElement(By.Name("fldInterest"));
                interestTxt.Clear();
                interestTxt.SendKeys(interestInt.ToString());
            }
            WaitFor(webDriver, 1);
            if (!string.IsNullOrEmpty(lateFee))
            {
                var feeInt = Math.Round(Convert.ToDecimal(lateFee), MidpointRounding.AwayFromZero);
                var feeTxt = webDriver.FindElement(By.Name("fldFee"));
                feeTxt.Clear();
                feeTxt.SendKeys(feeInt.ToString());
            }

            var nextBtn = webDriver.FindElements(By.XPath("//img[@src='gif/continue.gif']"))[0];
            nextBtn.Click();
            //WaitForReady(webDriver);
            WaitFor(webDriver, 2);
            var confirmBtn = webDriver.FindElements(By.XPath("//img[@src='gif/confirm.gif']"))[0];
            confirmBtn.Click();
            // WaitForReady(webDriver);

            WaitFor(webDriver, 3);
            service svc = new service();
            string otp = "";
            for (var i = 0; i < 120; i++)
            {
                var msg = svc.GetOTP(_bankLogin.LaneNo.Value);
                if (msg == null)
                {
                    Thread.Sleep(2000);
                    continue;
                }

                if (msg.Opt > 0)
                {
                    var txtOpt = msg.Opt.ToString();

                    otp = txtOpt.Length == 6 ? txtOpt : Regex.Match(msg.Body, "(\\d{6})").Groups[0].Value; ;
                    break;
                }

            }

            //Ask OTP if not received
            if (otp == "")
            {
                MessageBoxResult done = MessageBox.Show("Please fill the OTP and press OK button.", "Confirmation",
                                                        MessageBoxButton.OK, MessageBoxImage.Asterisk,
                                                        MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                //fill opt
                var optElm = webDriver.FindElement(By.Name("fldOtpToken"));
                optElm.SendKeys(otp);
            }

            //Delete otp
            svc.DeleteOTP(_bankLogin.LaneNo.Value);

            ///gif/submit.gif
            ///
            var submitBtn = webDriver.FindElements(By.XPath("//html/body/form[2]/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table[2]/tbody/tr[3]/td[3]/a"));
            submitBtn[0].Click();
            //WaitForReady(webDriver);
            WaitFor(webDriver, 2);

            var downloadBtn = webDriver.FindElements(By.XPath("//img[@src='gif/download.gif']"))[0];
            downloadBtn.Click();
            WaitFor(webDriver, 3);
        }

        private static void AssignAmount(IWebDriver webDriver, string amount)
        {
            if (amount.Length == 1)
            {
                var ones = webDriver.FindElement(By.Name("Ones"));
                var onesDDl = new SelectElement(ones);
                onesDDl.SelectByText(amount);
            }
            else if (amount.Length == 2)
            {
                var ones = webDriver.FindElement(By.Name("Ones"));
                var onesDDl = new SelectElement(ones);
                onesDDl.SelectByText(amount.Substring(1, 1));

                var ten = webDriver.FindElement(By.Name("Tens"));
                var tenDDl = new SelectElement(ten);
                tenDDl.SelectByText(amount.Substring(0, 1));
            }
            else if (amount.Length == 3)
            {
                var ones = webDriver.FindElement(By.Name("Ones"));
                var onesDDl = new SelectElement(ones);
                onesDDl.SelectByText(amount.Substring(2, 1));

                var ten = webDriver.FindElement(By.Name("Tens"));
                var tenDDl = new SelectElement(ten);
                tenDDl.SelectByText(amount.Substring(1, 1));

                var hundreds = webDriver.FindElement(By.Name("Hundreds"));
                var hundredsDDl = new SelectElement(hundreds);
                hundredsDDl.SelectByText(amount.Substring(0, 1));
            }
            else if (amount.Length > 3 && amount.Length < 6)
            {
                var ones = webDriver.FindElement(By.Name("Ones"));
                var onesDDl = new SelectElement(ones);
                onesDDl.SelectByText(amount.Substring(3, 1));

                var ten = webDriver.FindElement(By.Name("Tens"));
                var tenDDl = new SelectElement(ten);
                tenDDl.SelectByText(amount.Substring(2, 1));

                var hundreds = webDriver.FindElement(By.Name("Hundreds"));
                var hundredsDDl = new SelectElement(hundreds);
                hundredsDDl.SelectByText(amount.Substring(1, 1));

                var lngth = amount.Length == 4 ? 1 : 2;
                var thousands = webDriver.FindElement(By.Name("Thousands"));
                var thousandsDDl = new SelectElement(thousands);
                thousandsDDl.SelectByText(amount.Substring(0, lngth));

            }
            else if (amount.Length > 5 && amount.Length < 8)
            {
                var ones = webDriver.FindElement(By.Name("Ones"));
                var onesDDl = new SelectElement(ones);
                onesDDl.SelectByText(amount.Substring(4, 1));

                var ten = webDriver.FindElement(By.Name("Tens"));
                var tenDDl = new SelectElement(ten);
                tenDDl.SelectByText(amount.Substring(3, 1));

                var hundreds = webDriver.FindElement(By.Name("Hundreds"));
                var hundredsDDl = new SelectElement(hundreds);
                hundredsDDl.SelectByText(amount.Substring(2, 1));


                var thousands = webDriver.FindElement(By.Name("Thousands"));
                var thousandsDDl = new SelectElement(thousands);
                thousandsDDl.SelectByText(amount.Substring(1, 2));

                var lngth = amount.Length == 6 ? 1 : 2;
                var lakh = webDriver.FindElement(By.Name("Lakh"));
                var lakhDDl = new SelectElement(lakh);
                lakhDDl.SelectByText(amount.Substring(0, lngth));
            }

            else if (amount.Length > 7)
            {
                var ones = webDriver.FindElement(By.Name("Ones"));
                var onesDDl = new SelectElement(ones);
                onesDDl.SelectByText(amount.Substring(amount.Length - (amount.Length - 1), 1));

                var ten = webDriver.FindElement(By.Name("Tens"));
                var tenDDl = new SelectElement(ten);
                tenDDl.SelectByText(amount.Substring(amount.Length - (amount.Length - 2), 1));

                var hundreds = webDriver.FindElement(By.Name("Hundreds"));
                var hundredsDDl = new SelectElement(hundreds);
                hundredsDDl.SelectByText(amount.Substring(amount.Length - (amount.Length - 3), 1));


                var thousands = webDriver.FindElement(By.Name("Thousands"));
                var thousandsDDl = new SelectElement(thousands);
                thousandsDDl.SelectByText(amount.Substring(amount.Length - (amount.Length - 5), 2));


                var lakh = webDriver.FindElement(By.Name("Lakh"));
                var lakhDDl = new SelectElement(lakh);
                lakhDDl.SelectByText(amount.Substring(amount.Length - 7, 2));

                var lngth = amount.Length - 7;
                var crores = webDriver.FindElement(By.Name("Crores"));
                var croresDDl = new SelectElement(crores);
                croresDDl.SelectByText(amount.Substring(0, lngth));
            }

        }

        private static void ProcessGridData(IWebDriver webDriver)
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

            var gridElms = webDriver.FindElements(By.ClassName("gridauth_input_cell_style"));
            var firstLetter = gridElms[0].Text;
            var secondLetter = gridElms[1].Text;
            var thirdLetter = gridElms[2].Text;

            var firstInput = webDriver.FindElement(By.Id("TranRequestManagerFG.GRID_CARD_AUTH_VALUE_1__"));
            firstInput.SendKeys(grid[firstLetter]);
            var secondInput = webDriver.FindElement(By.Id("TranRequestManagerFG.GRID_CARD_AUTH_VALUE_2__"));
            secondInput.SendKeys(grid[secondLetter]);
            var thirstInput = webDriver.FindElement(By.Id("TranRequestManagerFG.GRID_CARD_AUTH_VALUE_3__"));
            thirstInput.SendKeys(grid[thirdLetter]);

        }
    }
}
