
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AutoFill
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Challan : Window
    {
        private service svc;
        private int transID;
        private decimal challanAmt;
        private RemittanceDto remittance;
        private UnzipFile unzipFile;
        private MultipartFormDataContent formData;
        private bool isFileBrowsed;
        public Challan(int trnasactionID,decimal challanAmount)
        {
            InitializeComponent();
            transID = trnasactionID;
            challanAmt = challanAmount;
            svc = new service();
            unzipFile = new UnzipFile();
            formData = null;
            LoadRemitance();
            ChallanProgressbar.Visibility = Visibility.Hidden;
        }

       private void LoadRemitance() {
            remittance = svc.GetRemitanceByTransID(transID);
            if (remittance.RemittanceID == 0)
            {
                ChallanAmount.Text = challanAmt.ToString();
                ChallanDate.Text = DateTime.Now.Date.ToString();
              //  upload.Visibility = Visibility.Hidden;
            }
            else {
               // upload.Visibility = Visibility.Visible;
                ChallanDate.Text = remittance.ChallanDate.ToString();
                ChallanNo.Text = remittance.ChallanID;
                AknowledgementNo.Text = remittance.ChallanAckNo;
                ChallanAmount.Text = remittance.ChallanAmount.ToString();
                // CustomerPropertyFileDto customerPropertyFileDto= svc.GetFile(remittance.ChallanFileID.ToString());
                CustomerPropertyFileDto customerPropertyFileDto = svc.GetFile(remittance.ChallanBlobID.ToString());
                if (customerPropertyFileDto != null)
                {
                    FileNameLabel.Content = customerPropertyFileDto.FileName;
                    upload.IsEnabled = false;
                }
            }
        }

        private void Upload_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            Nullable<bool> result = openFileDlg.ShowDialog();
            var filePath = openFileDlg.FileName;
            var challanDet = unzipFile.getChallanDetails(filePath,remittance.CustomerPAN);
            if (challanDet.Count == 0) {
                MessageBox.Show("PAN is not matched with uploaded file");
                return;
            }

            ChallanNo.Text = challanDet["serialNo"];
            AknowledgementNo.Text = challanDet["acknowledge"];
            FileNameLabel.Content = openFileDlg.SafeFileName;
            ChallanDate.Text = DateTime.ParseExact(challanDet["tenderDate"], "ddMMyy", null).ToString();

            remittance.ChallanIncomeTaxAmount = Convert.ToDecimal(challanDet["incomeTax"]);
            remittance.ChallanInterestAmount = Convert.ToDecimal(challanDet["interest"]);
            remittance.ChallanFeeAmount = Convert.ToDecimal(challanDet["fee"]);
            remittance.ChallanCustomerName = challanDet["name"].ToString();

          //var challanAmount = Convert.ToInt32(challanAmt);
            var challanAmount = Convert.ToDecimal(challanDet["challanAmount"]);
            if (challanAmount!= challanAmt)
                MessageBox.Show("Challan Amount is not matching");

            if (result == true)
            {
                 formData = new MultipartFormDataContent();
                var fileContent = new ByteArrayContent(File.ReadAllBytes(openFileDlg.FileName));
                var fileType = System.IO.Path.GetExtension(openFileDlg.FileName);
                var contentType = svc.GetContentType(fileType);
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
                var name = System.IO.Path.GetFileName(openFileDlg.FileName);
                formData.Add(fileContent, "file", name);
                isFileBrowsed = true;
              //  var bloblId = svc.UploadFile(formData, remittance.RemittanceID.ToString(), 7);
            }
        }

        private async void Save_Click(object sender, RoutedEventArgs e)
        {
            if (Validate()) {
                if (remittance.ClientPaymentTransactionID == 0)
                    remittance.ClientPaymentTransactionID = transID;
                remittance.ChallanAmount = Convert.ToDecimal(ChallanAmount.Text.Trim());
                remittance.ChallanAckNo= AknowledgementNo.Text.Trim();
                remittance.ChallanDate = Convert.ToDateTime(ChallanDate.Text.Trim());
                remittance.ChallanID= ChallanNo.Text.Trim();
                remittance.RemittanceStatusID = 2;
              

                ChallanProgressbar.Visibility = Visibility.Visible;
                int result = 0;
                await Task.Run(() =>
                {
                     result = svc.SaveRemittance(remittance);

                    if (result != 0)
                    {
                        if (isFileBrowsed)
                            SaveFile(result);
                        
                        MessageBox.Show("Challan details are saved successfully");                       
                    }
                    else
                        MessageBox.Show("Challan details are not saved ");

                });
                if (result != 0)
                { 
                    LoadRemitance(); 
                }
                    ChallanProgressbar.Visibility = Visibility.Hidden;
                this.DialogResult = false;//auto close
            }
        }

        private void SaveFile(int remittanceID)
        {
            var bloblId = svc.UploadFile(formData, remittanceID.ToString(), 7);
            isFileBrowsed = false;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        private bool Validate() {

            string challanDate = ChallanDate.Text.Trim();
            string challanNo = ChallanNo.Text.Trim();
            string acknowledgement = AknowledgementNo.Text.Trim();
            string challanAmount = ChallanAmount.Text.Trim();
            string errorMsg = "";
            if (challanDate == "")
                errorMsg = "Please enter the challan date";
            else if(challanNo=="")
                errorMsg = "Please enter the challan Number";
            else if (acknowledgement == "")
                errorMsg = "Please enter the challan Acknowledgement Number";
            else if (challanAmount == "")
                errorMsg = "Please enter the challan Amount";

            if (errorMsg != "")
            {
                MessageBox.Show(errorMsg);
                return false;
            }

            return true;
        }
    }
}
