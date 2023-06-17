using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Windows;

namespace AutoFill
{
    public class service
    {
        private HttpClient client;
        public service()
        {
            client = new HttpClient();
           //  client.BaseAddress = new Uri("http://leansyshost-001-site3.itempurl.com/api/"); //repro Live


            client.BaseAddress = new Uri("https://localhost:44301/api/");

            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public IList<TdsRemittanceDto> GetTdsRemitance(string custName,string premises,string unit, string fromUnit, string toUnit,string lot)
        {
            IList<TdsRemittanceDto> remitance = null;
            HttpResponseMessage response = new HttpResponseMessage();

            var query = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(custName))
                query["customerName"] = custName;
            if (!string.IsNullOrEmpty(premises))
                query["PropertyPremises"] = premises;
            if (!string.IsNullOrEmpty(unit))
                query["unitNo"] = unit;
            if (!string.IsNullOrEmpty(lot))
                query["lotNo"] = lot;

            if (!string.IsNullOrEmpty(fromUnit))
                query["fromUnitNo"] = fromUnit;
            if (!string.IsNullOrEmpty(toUnit))
                query["toUnitNo"] = toUnit;
            if (!string.IsNullOrEmpty(lot))
                query["lotNo"] = lot;
            //if (!string.IsNullOrEmpty(remittanceStatusID))
            //    query["remittanceStatusID"] = remittanceStatusID;

            response = client.GetAsync(QueryHelpers.AddQueryString("TdsRemittance/pendingTds", query)).Result;
           
            if (response.IsSuccessStatusCode)
            {
                remitance = response.Content.ReadAsAsync<IList<TdsRemittanceDto>>().Result;
            }
            return remitance;
        }

        public TdsRemittanceDto GetTdsRemitanceById(int clientPaymentTransactionID)
        {
            TdsRemittanceDto remitance = null;
            HttpResponseMessage response = new HttpResponseMessage();

            response = client.GetAsync("traces/"+ clientPaymentTransactionID).Result;

            if (response.IsSuccessStatusCode)
            {
                remitance = response.Content.ReadAsAsync<TdsRemittanceDto>().Result;
            }
            return remitance;
        }

        public string GetSellerPanByTransId(int clientPaymentTransactionID)
        {
            string pan = null;
            HttpResponseMessage response = new HttpResponseMessage();

            response = client.GetAsync("TdsRemittance/getSellerPan" + clientPaymentTransactionID).Result;

            if (response.IsSuccessStatusCode)
            {
                pan = response.Content.ReadAsAsync<string>().Result;
            }
            return pan;
        }

        public IList<RemittanceStatus> GetTdsRemitanceStatus()
        {
            IList<RemittanceStatus> remitance = null;
            HttpResponseMessage response = new HttpResponseMessage();

            response = client.GetAsync("RemittanceStatus").Result;

            if (response.IsSuccessStatusCode)
            {
                remitance = response.Content.ReadAsAsync<IList<RemittanceStatus>>().Result;
            }
            return remitance;
        }

        public IList<TdsRemittanceDto> GetTdsPaidList(string custName, string premises, string unit, string fromUnit, string toUnit, string lot, string remittanceStatusID)
        {
            IList<TdsRemittanceDto> remitance = null;
            HttpResponseMessage response = new HttpResponseMessage();

            var query = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(custName))
                query["customerName"] = custName;
            if (!string.IsNullOrEmpty(premises))
                query["PropertyPremises"] = premises;
            if (!string.IsNullOrEmpty(unit))
                query["unitNo"] = unit;
            if (!string.IsNullOrEmpty(fromUnit))
                query["fromUnitNo"] = fromUnit;
            if (!string.IsNullOrEmpty(toUnit))
                query["toUnitNo"] = toUnit;
            if (!string.IsNullOrEmpty(lot))
                query["lotNo"] = lot;
            if (!string.IsNullOrEmpty(remittanceStatusID))
                query["remittanceStatusID"] = remittanceStatusID;

            response = client.GetAsync(QueryHelpers.AddQueryString("TdsRemittance/processedList", query)).Result;

            if (response.IsSuccessStatusCode)
            {
                remitance = response.Content.ReadAsAsync<IList<TdsRemittanceDto>>().Result;
            }
            foreach(var entity in remitance){

                entity.OnlyTDS = !entity.OnlyTDS;
            }
            return remitance;
        }

        public AutoFillDto GetAutoFillData(int clientPaymentTransactionID)
        {

            AutoFillDto autoFillDto = null;
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("AutoFill/"+ clientPaymentTransactionID).Result;
          
            if (response.IsSuccessStatusCode)
            {
                autoFillDto = response.Content.ReadAsAsync<AutoFillDto>().Result;
            }
            return autoFillDto;
        }

        public BankAccountDetailsDto GetBankLoginDetails()
        {

            BankAccountDetailsDto bankDetail = null;
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("AutoFill/UserDetail").Result;

            if (response.IsSuccessStatusCode)
            {
                bankDetail = response.Content.ReadAsAsync<BankAccountDetailsDto>().Result;
            }
            return bankDetail;
        }

        public List<BankAccountDetailsDto> GetBankLoginList()
        {

            List<BankAccountDetailsDto> bankDetail = null;
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("AutoFill/AccountList").Result;

            if (response.IsSuccessStatusCode)
            {
                bankDetail = response.Content.ReadAsAsync<List<BankAccountDetailsDto>>().Result;
            }
            return bankDetail;
        }

        public bool SetToTdsPaid(int clientPaymentTransactionID)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.PutAsync("ClientPayment/remittanceStatus/" + clientPaymentTransactionID+"/2",null).Result;

            if (response.IsSuccessStatusCode)
            {
                return true;
            }
            return false;
        }

        public RemittanceDto GetRemitanceByTransID(int clientPaymentTransactionID)
        {
            RemittanceDto remittanceDto = null;
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("TdsRemittance/getRemittance/" + clientPaymentTransactionID).Result;

            if (response.IsSuccessStatusCode)
            {
                remittanceDto = response.Content.ReadAsAsync<RemittanceDto>().Result;
            }
            return remittanceDto;
        }

        public int SaveRemittance(RemittanceDto remittanceDto) {
            HttpResponseMessage response = new HttpResponseMessage();
            bool isNew = remittanceDto.RemittanceID == 0;

            //var json = JsonConvert.SerializeObject(remittanceDto);
            //var data = new StringContent(json, Encoding.UTF8, "application/json");

           CreateRemittaneCommand createRemittaneCommand=new CreateRemittaneCommand();
            createRemittaneCommand.remittanceDto = remittanceDto;

            if (isNew)
                response = client.PostAsJsonAsync("TdsRemittance" , createRemittaneCommand).Result;
            else
                response = client.PutAsJsonAsync("TdsRemittance", createRemittaneCommand).Result;

            if (response.IsSuccessStatusCode)
            {
                if (isNew)
                return response.Content.ReadAsAsync<int>().Result;

                return remittanceDto.RemittanceID;
            }
            return 0;
        }

        public int SaveDebitAdviceFile(MultipartFormDataContent file)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            response = client.PostAsync("DebitAdvice/uploadFile",  file).Result;

            if (response.IsSuccessStatusCode)
            {
                return response.Content.ReadAsAsync<int>().Result;

            }
            return 0;
        }

        public int SaveDebitAdvice(DebitAdviceDto debitAdviceDto)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            CreateDebitAdviceCommand debitAdviceObj = new CreateDebitAdviceCommand();
            debitAdviceObj.debitAdviceDto = debitAdviceDto;
            response = client.PostAsJsonAsync("DebitAdvice", debitAdviceObj).Result;

            if (response.IsSuccessStatusCode)
            {
                return response.Content.ReadAsAsync<int>().Result;

            }
            return 0;
        }

        public DebitAdviceDto GetDebitAdviceByClienttransId(int transId)
        {
            HttpResponseMessage response = new HttpResponseMessage();

            CreateDebitAdviceCommand debitAdviceObj = new CreateDebitAdviceCommand();
            response = client.GetAsync("DebitAdvice/getById/"+transId).Result;

            if (response.IsSuccessStatusCode)
            {
                return response.Content.ReadAsAsync<DebitAdviceDto>().Result;

            }

            return new DebitAdviceDto();
        }

        public string UploadFile(MultipartFormDataContent file,string remittanceId, int category) {
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.PostAsync("traces/"+ remittanceId + "/"+category, file).Result;
            return response.Content.ToString();
        }

        public CustomerPropertyFileDto GetFile(string blobId) {
            CustomerPropertyFileDto customerPropertyFileDto = null;
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("files/fileinfo/" + blobId).Result;
            if (response.IsSuccessStatusCode)
            {
                customerPropertyFileDto = response.Content.ReadAsAsync<CustomerPropertyFileDto>().Result;
            }
            return customerPropertyFileDto;

            //if (response.IsSuccessStatusCode)
            //{
            //    var ms = response.Content.ReadAsStreamAsync();
            //    var fs = File.Create()
            //}
        }

        public async  void DownloadFile(CustomerPropertyFileDto customerPropertyFileDto)
        {
            try
            {
                var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

                downloadPath += @"\REproFiles\" + customerPropertyFileDto.FileName;
                var ms = new MemoryStream(customerPropertyFileDto.FileBlob);
                var fs = File.Create(downloadPath);
                ms.Seek(0, SeekOrigin.Begin);
                ms.CopyTo(fs);
                MessageBox.Show("File is downloaded successfully. Please Refer the path : "+ downloadPath);
            }
            catch (Exception ex) {
                MessageBox.Show("File is not downloaded");
            }

            //    HttpResponseMessage response = new HttpResponseMessage();
            //    response = client.GetAsync("files/blobId/" + blobId).Result;

            //    if (response.IsSuccessStatusCode)
            //    {
            //        await using var ms = await response.Content.ReadAsStreamAsync();
            //        await using var fs = File.Create(fileName);
            //        ms.Seek(0, SeekOrigin.Begin);
            //        ms.CopyTo(fs);
            //    }
        }

        public bool DeleteRemittance(int remiitanceID)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.DeleteAsync("traces/" + remiitanceID).Result;

            if (response.IsSuccessStatusCode)
            {
                return true;

            }
            return false;
        }

        public bool SendMail(int transID)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("traces/sendmail/" + transID).Result;

            if (response.IsSuccessStatusCode)
            {
                return true;

            }
            return false;
        }
        public MessageDto GetOTP(int lane)
        {
            MessageDto msg = null;
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.GetAsync("Message/" + lane).Result;

            if (response.IsSuccessStatusCode)
            {
                msg = response.Content.ReadAsAsync<MessageDto>().Result;
            }
            return msg;
        }

        public bool DeleteOTP(int lane)
        {
            HttpResponseMessage response = new HttpResponseMessage();
            response = client.DeleteAsync("Message/" + lane).Result;

            if (response.IsSuccessStatusCode)
            {
                return true;
            }
            return false;
        }

        public string GetContentType(string fileExtension)
        {
            var mimeTypes = new Dictionary<String, String>
            {
                {".bmp", "image/bmp"},
                {".gif", "image/gif"},
                {".jpeg", "image/jpeg"},
                {".jpg", "image/jpeg"},
                {".png", "image/png"},
                {".tif", "image/tiff"},
                {".tiff", "image/tiff"},
                {".doc", "application/msword"},
                {".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                {".pdf", "application/pdf"},
                {".ppt", "application/vnd.ms-powerpoint"},
                {".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"},
                {".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                {".xls", "application/vnd.ms-excel"},
                {".csv", "text/csv"},
                {".xml", "text/xml"},
                {".txt", "text/plain"},
                {".zip", "application/zip"},
                {".ogg", "application/ogg"},
                {".mp3", "audio/mpeg"},
                {".wma", "audio/x-ms-wma"},
                {".wav", "audio/x-wav"},
                {".wmv", "audio/x-ms-wmv"},
                {".swf", "application/x-shockwave-flash"},
                {".avi", "video/avi"},
                {".mp4", "video/mp4"},
                {".mpeg", "video/mpeg"},
                {".mpg", "video/mpeg"},
                {".qt", "video/quicktime"}
            };

            // if the file type is not recognized, return "application/octet-stream" so the browser will simply download it
            return mimeTypes.ContainsKey(fileExtension) ? mimeTypes[fileExtension] : "application/octet-stream";
        }
    }

    public class BankAccountDetailsDto {
        public int AccountId { get; set; }
        public string UserName { get; set; }
        public string UserPassword { get; set; }
        public string BankName { get; set; }
        public int? LaneNo { get; set; }
        public string LetterA { get; set; }
        public string LetterB { get; set; }
        public string LetterC { get; set; }
        public string LetterD { get; set; }
        public string LetterE { get; set; }
        public string LetterF { get; set; }
        public string LetterG { get; set; }
        public string LetterH { get; set; }
        public string LetterI { get; set; }
        public string LetterJ { get; set; }
        public string LetterK { get; set; }
        public string LetterL { get; set; }
        public string LetterM { get; set; }
        public string LetterN { get; set; }
        public string LetterO { get; set; }
        public string LetterP { get; set; }
    }

    public class RemittanceStatus
    {        
        public int RemittanceStatusID { get; set; }
        public string RemittanceStatusText { get; set; }

    }
    public class BankList
    {
        public int BankID { get; set; }
        public string BankName { get; set; }

    }
    public class CustomerPropertyFileDto 
    {
        public int BlobID { get; set; }
        public Guid? OwnershipID { get; set; }
        public string FileName { get; set; }
        public byte[] FileBlob { get; set; }
        public DateTime? UploadTime { get; set; }

        public int? FileLength { get; set; }
        public string FileType { get; set; }
        public string PanID { get; set; }

        public int FileCategoryId { get; set; } = 4;
        
    }


    public class CreateRemittaneCommand {
        public RemittanceDto remittanceDto { get; set; }
    }

    public class RemittanceDto
    {
        public int RemittanceID { get; set; }
        public int ClientPaymentTransactionID { get; set; }
        public string ChallanID { get; set; }
        public DateTime ChallanDate { get; set; }
        public string ChallanAckNo { get; set; }
        public decimal ChallanAmount { get; set; }
        [ObsoleteAttribute]
        public Guid ChallanFileID { get; set; }
        public DateTime? F16BDateOfReq { get; set; }
        public string F16BRequestNo { get; set; }
        public string F16BCertificateNo { get; set; }
        [ObsoleteAttribute]
        public Guid F16BFileID { get; set; }
        public int RemittanceStatusID { get; set; }

        public int? Form16BlobID { get; set; }
        public int? ChallanBlobID { get; set; }
        public string F16CustName { get; set; }
        public DateTime? F16UpdateDate { get; set; }
        public decimal? F16CreditedAmount { get; set; }
        public bool? EmailSent { get; set; }
        public DateTime? EmailSentDate { get; set; }
        public decimal? ChallanIncomeTaxAmount { get; set; }
        public decimal? ChallanInterestAmount { get; set; }
        public decimal? ChallanFeeAmount { get; set; }
        public string ChallanCustomerName { get; set; }
        public virtual int UnitNo { get; set; }
        public virtual string CustomerName { get; set; }
        public virtual string Premises { get; set; }
        public virtual int LotNo { get; set; }
        public virtual DateTime DateOfBirth { get; set; }
        public virtual string CustomerPAN { get; set; }

    }

    public class TdsRemittanceDto
    {
        public bool IsSelected { get; set; }
        public int ClientPaymentTransactionID { get; set; }
        public int ClientPaymentID { get; set; }
        public Guid OwnershipID { get; set; }
        public string PropertyPremises { get; set; }
        public int UnitNo { get; set; }
        public bool TdsCollectedBySeller { get; set; }
        public Guid InstallmentID { get; set; }
        public DateTime? RevisedDateOfPayment { get; set; }
        public DateTime DateOfDeduction { get; set; }
        public string ReceiptNo { get; set; }
        public int LotNo { get; set; }
        public decimal AmountPaid { get; set; }
        public decimal SellerShare { get; set; }
        public string SellerName { get; set; }
        public string CustomerName { get; set; }
        public decimal CustomerShare { get; set; }
        public decimal GstAmount { get; set; }
        public decimal TdsAmount { get; set; }
        public decimal TdsInterest { get; set; }
        public decimal LateFee { get; set; }
        public decimal GrossAmount { get; set; }
        public decimal OwnershipAmount { get; set; }
        public int StatusTypeID { get; set; }
        public decimal GrossShareAmount { get; set; }
        public int RemittanceStatusID { get; set; }
        public string RemittanceStatus { get; set; }
        public decimal AmountShare { get; set; }

        public virtual DateTime ChallanDate { get; set; }
        public virtual string ChallanAckNo { get; set; }

        public virtual string CustomerPAN { get; set; }
        public virtual string SellerPAN { get; set; }
        public virtual string TracesPassword { get; set; }
        public virtual string AcknowledgementNo { get; set; }
        public virtual string AssessmentYear { get; set; }
        public DateTime? F16BDateOfReq { get; set; }
        public string F16BRequestNo { get; set; }
        public virtual decimal ChallanAmount { get; set; }

        public bool OnlyTDS { get; set; }
        public bool IsDebitAdvice { get; set; }
        public bool Show26qb { get; set; }
    }
    public class TracesModel : TdsRemittanceDto
    {
        public bool Form16BEnabled { get; set; }
    }
    public class AutoFillDto
    {
        public AutoFillDto()
        {
            tab1 = new Tab1();
            tab2 = new Tab2();
            tab3 = new Tab3();
            tab4 = new Tab4();
            eportal = new Eportal();
        }

        public Tab1 tab1 { get; set; }
        public Tab2 tab2 { get; set; }
        public Tab3 tab3 { get; set; }
        public Tab4 tab4 { get; set; }
        public Eportal eportal { get; set; }

    }

    public class Tab1
    {
        public string TaxApplicable { get; set; }

        public bool StatusOfPayee { get; set; }


        public string PanOfPayer { get; set; }
        public string PanOfTransferor { get; set; }

    }
    public class Tab2
    {
        //Transferee / buyer
        public string AddressPremisesOfTransferee { get; set; }
        public string AdressLine1OfTransferee { get; set; }
        public string AddressLine2OfTransferee { get; set; }
        public string CityOfTransferee { get; set; }
        public string StateOfTransferee { get; set; }
        public string PinCodeOfTransferee { get; set; }
        public string EmailOfOfTransferee { get; set; }
        public string MobileOfOfTransferee { get; set; }
        public bool IsCoTransferee { get; set; }



        //transferor/seller
        public string AddressPremisesOfTransferor { get; set; }
        public string AddressLine1OfTransferor { get; set; }
        public string AddressLine2OfTransferor { get; set; }
        public string CityOfTransferor { get; set; }
        public string StateOfTransferor { get; set; }
        public string PinCodeOfTransferor { get; set; }
        public string EmailOfOfTransferor { get; set; }
        public string MobileOfOfTransferor { get; set; }
        public bool IsCoTransferor { get; set; }

    }
    /// <summary>
    /// Property Details
    /// </summary>
    public class Tab3
    {
        //Property Details
        public string TypeOfProperty { get; set; }
        public string AddressPremisesOfProperty { get; set; }
        public string AddressLine1OfProperty { get; set; }
        public string AddressLine2OfProperty { get; set; }
        public string CityOfProperty { get; set; }
        public string StateOfProperty { get; set; }
        public string PinCodeOfProperty { get; set; }
        public DatePart DateOfAgreement { get; set; }

        public int TotalAmount { get; set; }
        public int PaymentType { get; set; }
        public DatePart RevisedDateOfPayment { get; set; }
        public DatePart DateOfDeduction { get; set; }

        public PlaceValues AmountPaidParts { get; set; }
        public int AmountPaid { get; set; }
        public Decimal BasicTax { get; set; }
        public Decimal Interest { get; set; }
        public Decimal LateFee { get; set; }
        public int StampDuty { get; set; }

        public Guid OwnershipId { get; set; }
        public Guid InstallmentId { get; set; }
        public int PropertyID { get; set; }
        public Decimal TotalAmountPaid { get; set; }
    }

    public class Tab4
    {

        //Payment Info
        public string ModeOfPayment { get; set; }
        public DatePart DateOfPayment { get; set; }
        public DatePart DateOfTaxDeduction { get; set; }
    }
    public class Eportal
    {
        public string LogInPan { get; set; }
        public string IncomeTaxPwd { get; set; }
        public bool IsCoOwners { get; set; }
        public string SellerPan { get; set; }
        public string SellerFlat { get; set; }
        public string SellerRoad { get; set; }
        public string SellerPinCode { get; set; }
        public string SellerPOstOffice { get; set; }
        public string SellerArea { get; set; }
        public string SellerMobile { get; set; }
        public string SellerEmail { get; set; }
        public bool IsLand { get; set; }
        public string PropFlat { get; set; }
        public string PropRoad { get; set; }
        public string PropPinCode { get; set; }
        public string PropPOstOffice { get; set; }
        public string PropArea { get; set; }
        public int paymentType { get; set; }
        public DatePart DateOfAgreement { get; set; }
        public int TotalAmount { get; set; }
        public int StampDuty { get; set; }
        public DatePart RevisedDateOfPayment { get; set; }
        public decimal TotalAmountPaid { get; set; }
        public Decimal Tds { get; set; }
        public Decimal Interest { get; set; }
        public Decimal Fee { get; set; }
        public int AmountPaid { get; set; }
    }
    public class DatePart
    {
        public int Day { get; set; }
        public string Month { get; set; }
        public int Year { get; set; }
    }

    public class PlaceValues
    {
        public int Crores { get; set; }
        public int Lakhs { get; set; }
        public int Thousands { get; set; }
        public int Hundreds { get; set; }
        public int Tens { get; set; }
        public int Ones { get; set; }
    }
    public class MessageDto
    {
        public int MessageID { get; set; }

        public string Subject { get; set; }
        public string Body { get; set; }
        public bool? Verified { get; set; }
        public int? Lane { get; set; }
        public string Message { get; set; }
        public int? Error_code { get; set; }
        public int Opt { get; set; }
    }

    public class CreateDebitAdviceCommand {
        public DebitAdviceDto debitAdviceDto { get; set; }
    }

    public class DebitAdviceDto
    {
        public int DebitAdviceID { get; set; }
        public int ClientPaymentTransactionID { get; set; }
        public string CinNo { get; set; }
        public DateTime? PaymentDate { get; set; }
        public int? BlobId { get; set; }
    }
}
