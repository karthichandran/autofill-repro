using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Windows;

namespace AutoFill
{
    public class AutoUploadService
    {
        private service svc;
        private UnzipFile unzipFile;
        public AutoUploadService()
        {
            unzipFile = new UnzipFile();
            svc = new service();
        }
        public bool UploadForm16b(string filePath, TdsRemittanceDto tdsRemittanceDto)
        {
            try
            {
                var form16bDetail = unzipFile.GetForm16bDetailsFromPDF(filePath, tdsRemittanceDto.CustomerPAN);

                MultipartFormDataContent formData = new MultipartFormDataContent();
                var fileContent = new ByteArrayContent(File.ReadAllBytes(filePath));
                var fileType = Path.GetExtension(filePath);
                var contentType = svc.GetContentType(fileType);
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
                var name = Path.GetFileName(filePath);
                formData.Add(fileContent, "file", name);

                var remittance = svc.GetRemitanceByTransID(tdsRemittanceDto.ClientPaymentTransactionID);
                remittance.F16BCertificateNo = form16bDetail["certNo"];
                remittance.F16CustName = form16bDetail["name"];
                if (remittance.F16BDateOfReq == null)
                    remittance.F16BDateOfReq = DateTime.Now.Date;

                if (remittance.F16BCertificateNo != "")
                    remittance.RemittanceStatusID = 4;
                else
                    remittance.RemittanceStatusID = 3;

                var updatDate = DateTime.ParseExact(form16bDetail["paymentDate"], "dd-MMM-yyyy", null);

                if (form16bDetail["amount"] != "")
                    remittance.F16CreditedAmount = Convert.ToDecimal(form16bDetail["amount"].Trim());
                if (updatDate != null)
                    remittance.F16UpdateDate = updatDate;

                var result = svc.SaveRemittance(remittance);
                if (result != 0)
                {
                    var bloblId = svc.UploadFile(formData, remittance.RemittanceID.ToString(), 6);
                    return true;
                    // MessageBox.Show("Request details are saved successfully");
                }
                else
                    return false;

            }
            catch (Exception ex) {
                return false;
                // MessageBox.Show("Update is failed.");
            }
        }

        public void UpdateForm16BRequestNo(int clientPaymentTransactionID, string reqNo) {
            try
            {
                var remittance = svc.GetRemitanceByTransID(clientPaymentTransactionID);
                remittance.F16BRequestNo = reqNo;
                if (remittance.F16BDateOfReq == null)
                    remittance.F16BDateOfReq = DateTime.Now.Date;

                if (remittance.F16BCertificateNo != "")
                    remittance.RemittanceStatusID = 4;
                else
                    remittance.RemittanceStatusID = 3;

                svc.SaveRemittance(remittance);
            }
            catch (Exception ex)
            {
               // MessageBox.Show("Update is failed.");
            }
        }
    }
}
