using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace AutoFill
{
    /// <summary>
    /// Interaction logic for Traces.xaml
    /// </summary>
    public partial class Traces : Window
    {
        private service svc;
       
        private decimal challanAmt;
        private string requestNo;
        private RemittanceDto remittance;
        private TdsRemittanceDto tdsRemittanceDto;
        private UnzipFile unzipFile;
        private MultipartFormDataContent formData;
        private bool isFileBrowsed;
        private CustomerPropertyFileDto customerPropertyFileDto;
        public Traces(TdsRemittanceDto model,string reqNo="")
        {
            InitializeComponent();
            tdsRemittanceDto = model;
             challanAmt = model.TdsAmount + model.TdsInterest + model.LateFee; 
            requestNo = reqNo;
            svc = new service();
            unzipFile = new UnzipFile();
            formData = null;
            LoadRemitance();
            TraceProgressbar.Visibility = Visibility.Hidden;
        }
        private void LoadRemitance()
        {
            remittance = svc.GetRemitanceByTransID(tdsRemittanceDto.ClientPaymentTransactionID);

            if (remittance.F16BDateOfReq == null)
                RequestDate.Text = DateTime.Now.Date.ToString();
            else
                RequestDate.Text = remittance.F16BDateOfReq.ToString();

          //  upload.Visibility = Visibility.Visible;

            customerPan.Text = remittance.CustomerPAN;
            dateOfBirth.Text = remittance.DateOfBirth.ToString("ddMMyyyy");
            if (remittance.F16BRequestNo != "")
                RequestNo.Text = remittance.F16BRequestNo;
            else
                RequestNo.Text = requestNo;

            CertificateNo.Text = remittance.F16BCertificateNo;

            if (remittance.F16CreditedAmount != null)
                PaidAmount.Text = remittance.F16CreditedAmount.ToString();
            if (remittance.F16CustName != null)
                CustomerName.Text = remittance.F16CustName;
            if (remittance.F16UpdateDate != null)
                UpdatedDate.Text = remittance.F16UpdateDate.Value.ToString("dd-MMM-yyyy");

                customerPropertyFileDto = svc.GetFile(remittance.Form16BlobID.ToString());
            if (customerPropertyFileDto != null)
            {
                FileNameLabel.Content = customerPropertyFileDto.FileName;
                //  upload.IsEnabled = false;
            }
            else {
                download.Visibility = Visibility.Hidden;
            }

        }

        private void Upload_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true)
            {
                FileNameLabel.Content = openFileDlg.SafeFileName;
                var filePath = openFileDlg.FileName;
                var revisedDate = tdsRemittanceDto.RevisedDateOfPayment;
               // var assessYear = revisedDate.Value.Year.ToString()+"-"+ revisedDate.Value.AddYears(1).ToString("yy");
                var form16bDetail = unzipFile.GetForm16bDetailsFromPDF(filePath, tdsRemittanceDto.CustomerPAN);
                CertificateNo.Text = form16bDetail["certNo"];
                CustomerName.Text = form16bDetail["name"];
                PaidAmount.Text= form16bDetail["amount"];
                var updatDate = DateTime.ParseExact(form16bDetail["paymentDate"], "dd-MMM-yyyy",null);
                UpdatedDate.Text = updatDate.ToString("dd-MMM-yyyy");

                formData = new MultipartFormDataContent();
                var fileContent = new ByteArrayContent(File.ReadAllBytes(openFileDlg.FileName));
                var fileType = System.IO.Path.GetExtension(openFileDlg.FileName);
                var contentType = svc.GetContentType(fileType);
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
                var name = System.IO.Path.GetFileName(openFileDlg.FileName);
                formData.Add(fileContent, "file", name);
                isFileBrowsed = true;
            }
        }

        private void Download_Click(object sender, RoutedEventArgs e)
        {
            if (customerPropertyFileDto != null)
            svc.DownloadFile(customerPropertyFileDto);

        }

        private async void Save_Click(object sender, RoutedEventArgs e)
        {
            if (Validate())
            {
                if (remittance.ClientPaymentTransactionID == 0)
                    remittance.ClientPaymentTransactionID = tdsRemittanceDto.ClientPaymentTransactionID;
                remittance.F16BCertificateNo = CertificateNo.Text.Trim();
                remittance.F16BDateOfReq = Convert.ToDateTime(RequestDate.Text.Trim());
                remittance.F16BRequestNo = RequestNo.Text.Trim();
               if( remittance.F16BCertificateNo!="")
                remittance.RemittanceStatusID = 4;
               else
                    remittance.RemittanceStatusID = 3;

                remittance.F16CustName = CustomerName.Text.Trim();
                if(PaidAmount.Text!="")
                remittance.F16CreditedAmount =Convert.ToDecimal( PaidAmount.Text.Trim());
                if(UpdatedDate.Text!="")
                    remittance.F16UpdateDate = Convert.ToDateTime(UpdatedDate.Text.Trim());

                TraceProgressbar.Visibility = Visibility.Visible;
                int result = 0;
                await Task.Run(() =>
                {
                    result = svc.SaveRemittance(remittance);
                if (result!=0)
                {
                    if (isFileBrowsed)
                        SaveFile();

                    MessageBox.Show("Request details are saved successfully");
                }
                else
                    MessageBox.Show("Request details are not saved ");
                });
                if (result != 0)
                {
                    LoadRemitance();
                }
                TraceProgressbar.Visibility = Visibility.Hidden;
            }
        }

        private void SaveFile() {
            var bloblId = svc.UploadFile(formData, remittance.RemittanceID.ToString(), 6);
            isFileBrowsed = false;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        private bool Validate()
        {

            string requestDate = RequestDate.Text.Trim();          
            string errorMsg = "";
            if (requestDate == "")
                errorMsg = "Please enter the request date";       
          
            if (errorMsg != "")
            {
                MessageBox.Show(errorMsg);
                return false;
            }

            return true;
        }
    }

}
