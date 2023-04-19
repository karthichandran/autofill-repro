using Aspose.Zip;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
//using iTextSharp.text.pdf;
//using iTextSharp.text.io;
using System.Text.RegularExpressions;
using System.Windows;

namespace AutoFill
{
    public class UnzipFile
    {
        public UnzipFile()
        {
        }

        public string extractFile(string fileName, string pwd)
        {
            // using Microsoft.Win32;

            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            var filePath = @downloadPath + "\\" + fileName + ".zip";
            downloadPath += @"\REproFiles";
            var startTime = DateTime.Now;
            while (!File.Exists(filePath))
            {
                Thread.Sleep(1000);
                var currentDate = DateTime.Now;
                if (currentDate.Subtract(startTime).TotalMinutes > 3)
                    break;
            }

            //// Open ZIP file
            using (FileStream zipFile = File.Open(filePath, FileMode.Open))
            {
                // Decrypt using password
                using (var archive = new Archive(zipFile, new ArchiveLoadOptions() { DecryptionPassword = pwd }))
                {
                    // Extract files to folder
                    archive.ExtractToDirectory(@downloadPath);
                }
            }
            //MessageBoxResult result = MessageBox.Show(String.Format("Form 16B with file name {0} downloaded successfully", fileName), "Confirmation",
            //                                         MessageBoxButton.OK, MessageBoxImage.Information,
            //                                         MessageBoxResult.OK, MessageBoxOptions.DefaultDesktopOnly);
            return @downloadPath + '\\' + fileName + ".pdf";
        }

        public Dictionary<string, string> getDebitAdviceDetails(string filePath)
        {
            Dictionary<string, string> challanDet = new Dictionary<string, string>();
           
            string text;
            using (var stream = File.OpenRead(filePath))
            using (UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream))
            {
                var page = document.GetPage(1);
                text = string.Join(" ", page.GetWords());
            }

            Console.WriteLine(text);
            var cinNo = GetWordAfterMatch(text, "CIN No");
            Console.WriteLine("Challan Serial NO :" + cinNo);
            var paymentDate = GetDate(text, "Time");
           
            challanDet.Add("cinNo", cinNo.ToString());
            challanDet.Add("paymentDate", paymentDate.ToString());

            return challanDet;
        }

        public Dictionary<string, string> getChallanDetails_da(string filePath)
        {
            Dictionary<string, string> challanDet = new Dictionary<string, string>();

            string text;
            using (var stream = File.OpenRead(filePath))
            using (UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream))
            {
                var page = document.GetPage(1);
                text = string.Join(" ", page.GetWords());
            }

            Console.WriteLine(text);
            var ackowledgeNo = GetWordAfterMatch(text, "Acknowledgement Number :");
            Console.WriteLine("Acknowledgement :" + ackowledgeNo);

            var serialNo = GetWordAfterMatch(text, "Challan No :");
            Console.WriteLine("Challan Serial NO :" + serialNo);

            var name = GetFullName(text);           

            var tenderDate = GetDate(text, "Tender Date :");

            var incomeTax = GetAmount(text);
           
            challanDet.Add("acknowledge", ackowledgeNo.ToString());
            challanDet.Add("serialNo", serialNo.ToString());
            challanDet.Add("name", name.ToString());
            challanDet.Add("tenderDate", tenderDate.ToString());
            challanDet.Add("amount", incomeTax.ToString());

            return challanDet;
        }

        public Dictionary<string, string> getChallanDetails(string filePath, string pan)
        {
            Dictionary<string, string> challanDet = new Dictionary<string, string>();
            //PDFParser pdfParser = new PDFParser();
            //PdfReader reader = new PdfReader(@filePath);
            //var text = new PDFParser().ExtractTextFromPDFBytes(reader.GetPageContent(1)).Trim().ToString();
            string text;
            using (var stream = File.OpenRead(filePath))
            using (UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream))
            {
                var page = document.GetPage(1);
                text = string.Join(" ", page.GetWords());
            }

            Console.WriteLine(text);
            var serialNo = GetWordAfterMatch(text, "Challan Serial No.");
            Console.WriteLine("Challan Serial NO :" + serialNo);
            var paninDoc = GetWordAfterMatch(text, "PAN:");
            if (pan != paninDoc.ToString())
                return challanDet;
            challanDet.Add("serialNo", serialNo.ToString());

            var name = GetName(text, "Full Name :");
            challanDet.Add("name", name.ToString());
            //var itns = GetWordAfterMatch(text, "Challan No./ITNS");
            //Console.WriteLine("ITNS :" + itns);
            var tenderDate = GetWordAfterMatch(text, "Tender Date");
            challanDet.Add("tenderDate", tenderDate.ToString());
            var challamAmount = GetWordAfterMatch(text, "Rs. :");
            challanDet.Add("challanAmount", challamAmount.ToString());
            // var PAN = "BUZPP5880P"; //todo pass the pan number
            // pan = "BBFPK5517D";
            var tds = GetTDSConfirmationNo(text, pan);
            Console.WriteLine("tds conf NO :" + tds);
            challanDet.Add("acknowledge", tds.ToString());
            Console.ReadLine();

            var incomeTax = GetWordAfterMatch(text, "Income Tax");
            challanDet.Add("incomeTax", incomeTax.ToString());
            var interest = GetWordAfterMatch(text, "Interest");
            challanDet.Add("interest", interest.ToString());
            var fee = GetWordAfterMatch(text, "Fee");
            challanDet.Add("fee", fee.ToString());
            return challanDet;
        }

        public Dictionary<string, string> getChallanDetails_Hdfc(string filePath, string sellerPan)
        {
            Dictionary<string, string> challanDet = new Dictionary<string, string>();
            //PDFParser pdfParser = new PDFParser();
            //PdfReader reader = new PdfReader(@filePath);
            //var text = new PDFParser().ExtractTextFromPDFBytes(reader.GetPageContent(1)).Trim().ToString();
            string text;
            using (var stream = File.OpenRead(filePath))
            using (UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream))
            {
                var page = document.GetPage(1);
                text = string.Join(" ", page.GetWords());
            }

            Console.WriteLine(text);
            var serialNo = GetWordAfterMatch(text, "ChallanSerialNo");
            Console.WriteLine("Challan Serial NO :" + serialNo);
            //var paninDoc = GetWordAfterMatch(text, "PAN:");
            //if (pan != paninDoc.ToString())
            //    return challanDet;
            challanDet.Add("serialNo", serialNo.ToString());

            var name = GetWordAfterMatch(text, "oftheAssessee");
            challanDet.Add("name", name.ToString());
            //var itns = GetWordAfterMatch(text, "Challan No./ITNS");
            //Console.WriteLine("ITNS :" + itns);
            var tenderDate = GetDate(text, "ofReceipt");
            challanDet.Add("tenderDate", tenderDate.ToString());
            var challamAmount = GetWordAfterMatch(text, "TOTAL");
            challanDet.Add("challanAmount", challamAmount.ToString());
            // var PAN = "BUZPP5880P"; //todo pass the pan number
            // pan = "BBFPK5517D";
            var tds = GetWordAfterMatch(text, sellerPan);
            Console.WriteLine("tds conf NO :" + tds);
            challanDet.Add("acknowledge", tds.ToString());
            Console.ReadLine();

            var incomeTax = GetWordAfterMatch(text, "BasicTax");
            challanDet.Add("incomeTax", incomeTax.ToString());
            var interest = GetWordAfterMatch(text, "Interest");
            challanDet.Add("interest", interest.ToString());
            var fee = GetWordAfterMatch(text, "Fee");
            challanDet.Add("fee", fee.ToString());
            return challanDet;
        }

        private object GetDate(string text, string word)
        {
            var pattern = string.Format(@"\b\w*" + word + @"\s\w*(/)\w+(/)\w+\b");
            string match = Regex.Match(text, @pattern).Groups[0].Value;
            string[] words = match.Split(' ');
            string wordAfter = words[words.Length - 1];

            return wordAfter;
        }

        private object GetName(string text, string word)
        {
            var pattern = string.Format(@"\b\w*" + word + @"\w*\s+[^0-9]*");
            string match = Regex.Match(text, @pattern).Groups[0].Value;
            string[] words = match.Split(':');
            string wordAfter = words[words.Length - 1];

            return wordAfter;
        }
        private object GetFullName(string text)
        {
            string ward = Regex.Match(text, "Name : (.*)Assessment").Groups[1].Value;      
            return ward;
        }
        private object GetAmount(string text)
        {
            string ward = Regex.Match(text, "₹(.*)Amount").Groups[1].Value;
            return ward;
        }

        private object GetWordAfterMatch(string text, string word)
        {

            var pattern = string.Format(@"\b\w*" + word + @"\w*\s+\w+\b");
            string match = Regex.Match(text, @pattern).Groups[0].Value;
            string[] words = match.Split(' ');
            string wordAfter = words[words.Length - 1];

            return wordAfter;
        }

        private object GetTDSConfirmationNo(string text, string word)
        {

            var pattern = string.Format(word + @"?.*");
            string match = Regex.Match(text, @pattern).Groups[0].Value.Substring(25, 100);
            string[] words = match.Split(',');
            string wordAfter = words[3];

            return wordAfter;
        }

        private object GetCertificateNoAfterMatch(string text, string word)
        {
            var pattern = string.Format(@"\b\w*" + word + @"\w*\s+\w+\s+\w+(-)\w+\s+\w+\b");
            string match = Regex.Match(text, @pattern).Groups[0].Value;
            string[] words = match.Split(' ');
            string wordAfter = words[words.Length - 1];
            return wordAfter;
        }

        public Dictionary<string, string> GetForm16bDetailsFromPDF(string filePath, string pan)
        {
            // pan = "AMSPA9519Q";
            Dictionary<string, string> form16bDet = new Dictionary<string, string>();
            //PDFParser pdfParser = new PDFParser();
            //PdfReader reader = new PdfReader(@filePath);
            //var text = new PDFParser().ExtractTextFromPDFBytes(reader.GetPageContent(1)).Trim().ToString();
            //Console.WriteLine(text);
            string text;
            using (var stream = File.OpenRead(filePath))
            using (UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(stream))
            {
                var page = document.GetPage(1);
                text = string.Join(" ", page.GetWords());
            }
            //  var certNo = GetCertificateNoAfterMatch(text, pan);
            var certNo = GetWordAfterMatch(text, "Certificate No.:");
            form16bDet.Add("certNo", certNo.ToString());

            // var datePattern = string.Format(@"\b\w*" + pan + @"\w*\s+\w+\s+\w+(-)\w+\s+\w+\s+\w+(-)\w+(-)\w+\b");            
            //string match = Regex.Match(text, @datePattern).Groups[0].Value;
            //string[] dateArry = match.Split(' ');
            //string date = dateArry[dateArry.Length - 1];
            //form16bDet.Add("paymentDate", date);
            var datePattern = string.Format(@"\b\w*" + "Updated On:" + @"\s\w*(-)\w+(-)\w+\b");
            string match = Regex.Match(text, @datePattern).Groups[0].Value;
            string[] dateArry = match.Split(':');
            string date = dateArry[dateArry.Length - 1];
            form16bDet.Add("paymentDate", date.Trim());

            //var namePattern = string.Format(@"\b\w*" + pan + @"\w*\s+\w+\s+\w+(-)\w+\s+\w+\s+\w+(-)\w+(-)\w+[\s+\w+]*,");           
            //string nameMatch = Regex.Match(text, @namePattern).Groups[0].Value;
            //string[] nameArray = nameMatch.Split(' ');
            //string name = "";
            //int inx = nameArray.Length - 5;
            //for (int i = 0; i < inx-1; i++)
            //{
            //    name += nameArray[5 + i] + " ";
            //}
            //form16bDet.Add("name", name.Split(',')[0]);

            var namePattern = string.Format(@"\b\w*" + "Full Name:" + @"(.*)");
            string nameMatch = Regex.Match(text, @namePattern).Groups[1].Value;
            string[] nameArray = nameMatch.Split(new string[] { "Page" }, StringSplitOptions.None);
            form16bDet.Add("name", nameArray[0].Trim());

            //var amountPattern = string.Format(@"\b\w*sum of Rs.\w*\s+\w*.\w*");
            //string amountMatch = Regex.Match(text, @amountPattern).Groups[0].Value;
            //string[] amountArry = amountMatch.Split(' ');
            //string amount = amountArry[amountArry.Length-1];
            //amount = amount.Substring(3, amount.Length - 3);
            //form16bDet.Add("amount", amount);

            var amountPattern = string.Format(@"\b\w*sum of Rs.\w*\s+\w*.\w*");
            string amountMatch = Regex.Match(text, @amountPattern).Groups[0].Value;
            string[] amountArry = amountMatch.Split(' ');
            string amount = amountArry[amountArry.Length - 1];
            form16bDet.Add("amount", amount);



            return form16bDet;
        }


    }
}
