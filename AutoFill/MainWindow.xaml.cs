using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AutoFill
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       private service svc;
        private string tds;
        private string tdsInterest;
        private string lateFee;
        private IList<RemittanceStatus> remittanceStatusList;
        private List<BankAccountDetailsDto> accountList;
        private BankAccountDetailsDto bankLogin;

        private int selectedAccount = 0;
        private string selectedBank = "";

        //filter fields
        private string remittanceStatusID;        
        private string custName; 
        private string premise;
        private string unit;
        private string fromUnit;
        private string toUnit;
        private string lot;

        ObservableCollection<TdsRemittanceDto> tdsRemitanceList { get; set; }

        BackgroundWorker worker;
        public MainWindow()
        {
            InitializeComponent();
            svc = new service();

            LoadBankAccountList();
            // LoadRemitance();
            LoadRemittanceStatus();
             progressbar1.Visibility = Visibility.Hidden;
            TracesProgressbar.Visibility = Visibility.Hidden;
        }
        private void LoadBankAccountList() {
            accountList = svc.GetBankLoginList();
            accountddl.ItemsSource = accountList;
            accountddl.DisplayMemberPath = "UserName";
            accountddl.SelectedValuePath = "AccountId";


        //    List<BankList> banks = new List<BankList>() { new BankList() { BankID = 1, BankName = "ICICI" }, new BankList() { BankID = 2, BankName = "HDFC" } };
        //    bankddl.ItemsSource = banks;
        //    bankddl.DisplayMemberPath = "BankName";
        //    bankddl.SelectedValuePath = "BankName";
        //    bankddl.SelectedIndex = 0;
        }

        private void LoadRemitance() {
            IList<TdsRemittanceDto> remitanceList = svc.GetTdsRemitance("","","","","","");
            remitanceList = remitanceList.OrderBy(x => x.UnitNo).ToList();
            remitanceGrid.ItemsSource = remitanceList;           

            TotalRecordsLbl.Content = remitanceList.Count;
            var totalTds = remitanceList.Sum(x => x.TdsAmount);
            TotalTDSLbl.Content = totalTds;

        }

        private void LoadRemittanceStatus() {
            remittanceStatusList = svc.GetTdsRemitanceStatus();
            var emptyObj = new RemittanceStatus() { RemittanceStatusText = "", RemittanceStatusID = -1 };
            remittanceStatusList.Insert(0, emptyObj);

            tracesRemitanceStatusddl.ItemsSource = remittanceStatusList;
            tracesRemitanceStatusddl.DisplayMemberPath = "RemittanceStatusText";
            tracesRemitanceStatusddl.SelectedValuePath = "RemittanceStatusID";
        }

        private bool AutoFillForm26Q(int clientPaymentTransactionID) {
            AutoFillDto autoFillDto = svc.GetAutoFillData(clientPaymentTransactionID);
            if (autoFillDto == null)
            {
                MessageBox.Show("Data is not available to proceed Form26QB", "alert", MessageBoxButton.OK);
                return false; 
            }

            if (selectedAccount == 0) {
                MessageBox.Show("Please select User Account", "alert", MessageBoxButton.OK);
                return false; 
            }


            var status = false;
            if (bankLogin.BankName == "HDFC")
                status = FillForm26Q.AutoFillForm26QB(autoFillDto, tds, tdsInterest, lateFee, bankLogin, clientPaymentTransactionID.ToString());
            else
                status = FillForm26QB_ICICI.AutoFillForm26QB(autoFillDto, tds, tdsInterest, lateFee, bankLogin, clientPaymentTransactionID.ToString());

            return status;
        }

        private void proceedForm(object sender, RoutedEventArgs e)
        {
            if (bankLogin == null) {
                MessageBox.Show("Please select account.");
                return;
            }
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            tds = model.TdsAmount.ToString();
            tdsInterest = model.TdsInterest.ToString();
            lateFee = model.LateFee.ToString();
            // MethodThatWillCallComObject(AutoFillForm26Q,model.ClientPaymentTransactionID);

            progressbar1.Visibility = Visibility.Visible;
           Task.Factory.StartNew(() =>
            {
                //fill form 26q then download
                var challanAmount = model.TdsAmount + model.TdsInterest + model.LateFee;
                var status = AutoFillForm26Q(model.ClientPaymentTransactionID);
                if (!status)
                    return;

                // auto upload
                // autoUploadChallan(model.ClientPaymentTransactionID, challanAmount, selectedBank,model.SellerPAN);
                UploadDebitAdvice(model.ClientPaymentTransactionID);

                // Reload filter
                this.Dispatcher.Invoke((Action)(() =>
                { 
                    var custName = customerNameTxt.Text;
                    var premise = PremisesTxt.Text;
                    var unit = unitNoTxt.Text;
                    var lot = lotNoTxt.Text;
                    var fromUnit = fromUnitNoTxt.Text;
                    var toUnit = toUnitNoTxt.Text;
                    var remitanceList= svc.GetTdsRemitance(custName, premise, unit, fromUnit, toUnit, lot);
                    remitanceList = remitanceList.OrderBy(x => x.UnitNo).ToList();
                    remitanceGrid.ItemsSource = remitanceList;
                    TotalRecordsLbl.Content = remitanceList.Count;
                    var totalTds = remitanceList.Sum(x => x.TdsAmount);
                    TotalTDSLbl.Content = totalTds;

                }));             
               

            }).ContinueWith(t =>
            {
                progressbar1.Visibility = Visibility.Hidden;
            },TaskScheduler.FromCurrentSynchronizationContext());         
           
        }

        private void BulkPayment_Click(object sender, RoutedEventArgs e)
        {
            
            var remittanceList = (List<TdsRemittanceDto>)remitanceGrid.ItemsSource;
            remittanceList = remittanceList.Where(x => x.IsSelected == true).ToList();
            if (remittanceList.Count() == 0)
                return;

            if (selectedAccount == 0)
            {
                MessageBox.Show("Please select User Account", "alert", MessageBoxButton.OK);
                return;
            }

            progressbar1.Visibility = Visibility.Visible;
            string failedPayments = "";
            Task.Factory.StartNew(() =>
            {
                foreach (var item in remittanceList)
                {
                    Dispatcher.BeginInvoke(new Action(() => lbl_runingUnit.Content = item.CustomerName + "  -  " + item.UnitNo + "  -  " + item.TdsAmount), System.Windows.Threading.DispatcherPriority.Send);

                    var challanAmount = item.TdsAmount + item.TdsInterest + item.LateFee;
                    var id = item.ClientPaymentTransactionID;
                    AutoFillDto autoFillDto = svc.GetAutoFillData(id);
                    if (autoFillDto == null)
                    {
                        continue;
                    }

                    var status = false;
                    if (bankLogin.BankName == "HDFC")
                        status = FillForm26Q.AutoFillForm26QB_NoMsg(autoFillDto, item.TdsAmount.ToString(), item.TdsInterest.ToString(), item.LateFee.ToString(), bankLogin, id.ToString());
                    else
                        status = FillForm26QB_ICICI.AutoFillForm26QB_NoMsg(autoFillDto, item.TdsAmount.ToString(), item.TdsInterest.ToString(), item.LateFee.ToString(), bankLogin, id.ToString());

                    if (!status)
                    {
                        if (failedPayments == "")
                        {
                            failedPayments = item.UnitNo + " - " + item.CustomerName;
                        }
                        else
                        {
                            failedPayments = failedPayments + " , " + item.UnitNo + " - " + item.CustomerName;
                        }
                        continue;
                    }

                    //autoUploadChallan_NoMsg(id, challanAmount, selectedBank,item.SellerPAN);    
                    UploadDebitAdvice(id);
                }


            }).ContinueWith(t =>
            {
                progressbar1.Visibility = Visibility.Hidden;
                RemittanceSearchFilter();
                if (failedPayments != "")
                {
                    MessageBox.Show("Following units are failed : " + failedPayments);
                }
                MessageBox.Show("Batch process completed.");
                Dispatcher.BeginInvoke(new Action(() => lbl_runingUnit.Content = ""), System.Windows.Threading.DispatcherPriority.Send);
            }, TaskScheduler.FromCurrentSynchronizationContext());

            

        }

        private void TdsPaid(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var challanAmount = model.TdsAmount + model.TdsInterest + model.LateFee;
            Challan challan = new Challan(model.ClientPaymentTransactionID, challanAmount);
            challan.Owner = this;
            challan.ShowDialog();
            RemittanceSearchFilter();
        }

        private void UploadDebitAdvice(int transID ) {
            var remittance = svc.GetRemitanceByTransID(transID);
            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            //var fileName = "AHUPB2786K_23040400148910ICIC_DTAX_04042023_TaxPayer.pdf";
            // var fileName = remittance.CustomerPAN + "_*.pdf";
            var fileName = "23040800002375ICIC_ChallanReceipt.pdf";

            var directory = new DirectoryInfo(downloadPath);
            var myFile = directory.GetFiles(fileName).OrderByDescending(f => f.LastWriteTime).ToList();

            if (myFile.Count == 0)
                return;

            var filename = myFile[0].FullName;

            var unzipFile = new UnzipFile();
            Dictionary<string, string> debitAdvice;

            debitAdvice = unzipFile.getDebitAdviceDetails(filename);          

            var formData = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(File.ReadAllBytes(filename));
            var fileType = System.IO.Path.GetExtension(filename);
            var contentType = svc.GetContentType(fileType);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
            var name = System.IO.Path.GetFileName(filename);
            formData.Add(fileContent, "file", name);
            int result = svc.SaveDebitAdviceFile(formData);

            if (result != 0)
            {
                DebitAdviceDto dto = new DebitAdviceDto
                {
                    ClientPaymentTransactionID = transID,
                    CinNo = debitAdvice["cinNo"],
                    PaymentDate = DateTime.ParseExact(debitAdvice["paymentDate"], "dd/MM/yyyy", null),
                    BlobId=result
                };
                var debitAdviceId = svc.SaveDebitAdvice(dto);
            }
        }

        private void UploadChallan(int transID)
        {
            //var remittance = svc.GetRemitanceByTransID(transID);
            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            //var fileName = "AHUPB2786K_23040400148910ICIC_DTAX_04042023_TaxPayer.pdf";
            // var fileName = remittance.CustomerPAN + "_*.pdf";
            var fileName = "23040800002375ICIC_ChallanReceipt.pdf";

            var directory = new DirectoryInfo(downloadPath);
            var myFile = directory.GetFiles(fileName).OrderByDescending(f => f.LastWriteTime).ToList();

            if (myFile.Count == 0)
                return;

            var filename = myFile[0].FullName;

            var unzipFile = new UnzipFile();
            Dictionary<string, string> debitAdvice;

            debitAdvice = unzipFile.getChallanDetails_da(filename);

           
        }

        private async void autoUploadChallan(int transID, decimal challanAmt, string bankName,string sellerPan)
        {
          
            var remittance = svc.GetRemitanceByTransID(transID);

            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();
            //  string[] filePaths = Directory.GetFiles(downloadPath, remittance.CustomerPAN + "_*.pdf").OrderByDescending(f=>f.LastWriteTime);

            var fileName = "";
            if (bankName == "HDFC")
                fileName = "*_*.pdf";
            else
                fileName = remittance.CustomerPAN + "_*.pdf";

            var directory = new DirectoryInfo(downloadPath);
            var myFile = directory.GetFiles(fileName).OrderByDescending(f => f.LastWriteTime).ToList();

            if (myFile.Count == 0)
                return;
          
            var filename = myFile[0].FullName;           

            var unzipFile = new UnzipFile();
            Dictionary<string, string> challanDet;
            if (bankName == "HDFC")
            {
               
                challanDet = unzipFile.getChallanDetails_Hdfc(filename, sellerPan);
            }
            else
                challanDet = unzipFile.getChallanDetails(filename, remittance.CustomerPAN);

            if (challanDet.Count == 0)
            {
                MessageBox.Show("PAN is not matched with uploaded file");
                return;
            }

            if (remittance.ClientPaymentTransactionID == 0)
                remittance.ClientPaymentTransactionID = transID;

            remittance.ChallanAmount = challanAmt;
            remittance.ChallanID = challanDet["serialNo"];
            remittance.ChallanAckNo = challanDet["acknowledge"];
            if (bankName == "HDFC")
                remittance.ChallanDate = DateTime.ParseExact(challanDet["tenderDate"], "dd/MM/yyyy", null);
            else
                remittance.ChallanDate = DateTime.ParseExact(challanDet["tenderDate"], "ddMMyy", null);
            remittance.RemittanceStatusID = 2;

            remittance.ChallanIncomeTaxAmount = Convert.ToDecimal(challanDet["incomeTax"]);
            remittance.ChallanInterestAmount = Convert.ToDecimal(challanDet["interest"]);
            remittance.ChallanFeeAmount = Convert.ToDecimal(challanDet["fee"]);
            remittance.ChallanCustomerName = challanDet["name"].ToString();

            //var challanAmount = Convert.ToDecimal(challanDet["challanAmount"]);
            //if (challanAmount != challanAmt)
            //    MessageBox.Show("Challan Amount is not matching");

            var formData = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(File.ReadAllBytes(filename));
            var fileType = System.IO.Path.GetExtension(filename);
            var contentType = svc.GetContentType(fileType);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
            var name = System.IO.Path.GetFileName(filename);
            formData.Add(fileContent, "file", name);

            int result = svc.SaveRemittance(remittance);

            if (result != 0)
            {
                var bloblId = svc.UploadFile(formData, result.ToString(), 7);

                MessageBox.Show("Challan details are saved successfully");
            }
            else
                MessageBox.Show("Challan details are not saved ");
        }

        private async void autoUploadChallan_NoMsg(int transID, decimal challanAmt, string bankName, string sellerPan)
        {
            var remittance = svc.GetRemitanceByTransID(transID);

            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();
            //  string[] filePaths = Directory.GetFiles(downloadPath, remittance.CustomerPAN + "_*.pdf").OrderByDescending(f=>f.LastWriteTime);

            var fileName = "";
            if (bankName == "HDFC")
                fileName = "*_*.pdf";
            else
                fileName = remittance.CustomerPAN + "_*.pdf";

            var directory = new DirectoryInfo(downloadPath);
            var myFile = directory.GetFiles(fileName).OrderByDescending(f => f.LastWriteTime).ToList();

            if (myFile.Count == 0)
                return;

            var filename = myFile[0].FullName;

            var unzipFile = new UnzipFile();
            Dictionary<string, string> challanDet;
            if (bankName == "HDFC")
            {
                challanDet = unzipFile.getChallanDetails_Hdfc(filename, sellerPan);
            }
            else
                challanDet = unzipFile.getChallanDetails(filename, remittance.CustomerPAN);

            if (challanDet.Count == 0)
                return;

            if (remittance.ClientPaymentTransactionID == 0)
                remittance.ClientPaymentTransactionID = transID;

            remittance.ChallanAmount = challanAmt;
            remittance.ChallanID = challanDet["serialNo"];
            remittance.ChallanAckNo = challanDet["acknowledge"];
            if (bankName == "HDFC")
                remittance.ChallanDate = DateTime.ParseExact(challanDet["tenderDate"], "dd/MM/yyyy", null);
            else
                remittance.ChallanDate = DateTime.ParseExact(challanDet["tenderDate"], "ddMMyy", null);
            remittance.RemittanceStatusID = 2;

            remittance.ChallanIncomeTaxAmount = Convert.ToDecimal(challanDet["incomeTax"]);
            remittance.ChallanInterestAmount = Convert.ToDecimal(challanDet["interest"]);
            remittance.ChallanFeeAmount = Convert.ToDecimal(challanDet["fee"]);
            remittance.ChallanCustomerName = challanDet["name"].ToString();

            //var challanAmount = Convert.ToDecimal(challanDet["challanAmount"]);
            //if (challanAmount != challanAmt)
            //    MessageBox.Show("Challan Amount is not matching");

            var formData = new MultipartFormDataContent();
            var fileContent = new ByteArrayContent(File.ReadAllBytes(filename));
            var fileType = System.IO.Path.GetExtension(filename);
            var contentType = svc.GetContentType(fileType);
            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType);
            var name = System.IO.Path.GetFileName(filename);
            formData.Add(fileContent, "file", name);

            int result = svc.SaveRemittance(remittance);

            if (result != 0)
            {
                var bloblId = svc.UploadFile(formData, result.ToString(), 7);
            }

        }


        //private void MethodThatWillCallComObject(TdsRemittanceDto model)
        //{
        //    progressbar1.Visibility = Visibility.Visible;
        //    System.Threading.Tasks.Task.Factory.StartNew(() =>
        //    {
        //        //this will call in background thread               
        //        AutoFillForm296Q(model.ClientPaymentTransactionID);

        //    }).ContinueWith(t =>
        //    {
        //        progressbar1.Visibility = Visibility.Hidden;
        //    }, System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext());
        //}

        private void MethodThatWillCallComObject(Func<int,bool> function,int id)
        {
            progressbar1.Visibility = Visibility.Visible;
            System.Threading.Tasks.Task.Factory.StartNew(() =>
            {
                function(id);

            }).ContinueWith(t =>
            {
                progressbar1.Visibility = Visibility.Hidden;
            }, System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void Search_Click(object sender, RoutedEventArgs e)
        {
            // RemittanceSearchFilter();   
            // UploadDebitAdvice(5609);
            UploadChallan(5609);
        }

        private void TracesSearch_Click(object sender, RoutedEventArgs e)
        {
            TracesSearchFilter();
        }

        private void TracesReset_Click(object sender, RoutedEventArgs e)
        {
            tracesCustomerNameTxt.Text = "";
            tracesPremisesTxt.Text = "";
            tracesUnitNoTxt.Text = "";
            tracesFromUnitNoTxt.Text = "";
            tracesToUnitNoTxt.Text = "";
            tracesLotNoTxt.Text = "";

            tracesRemitanceStatusddl.SelectedValue = -1;
        }

        private void textboxKeydown(object sender, KeyEventArgs e)
        {
             if ( e.Key==Key.Enter)
                RemittanceSearchFilter();
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            customerNameTxt.Text = "";
            PremisesTxt.Text = "";
            unitNoTxt.Text = "";
            lotNoTxt.Text = "";
            fromUnitNoTxt.Text = "";
            toUnitNoTxt.Text = "";

            remitanceGrid.ItemsSource = null;
            TotalRecordsLbl.Content = 0;
            TotalTDSLbl.Content = 0;
        }
        private void RemittanceSearchFilter()
        {
            var custName = customerNameTxt.Text;
            var premise = PremisesTxt.Text;
            var unit = unitNoTxt.Text;
            var lot = lotNoTxt.Text;
            var fromUnit = fromUnitNoTxt.Text;
            var toUnit = toUnitNoTxt.Text;

            if (!ValidateRequiredFields()) {
                MessageBox.Show("Please fill the mandatory fields");
                return;
            }

            RemittanceSearchTask(custName, premise, unit,fromUnit,toUnit, lot);

        }

        private bool ValidateRequiredFields() {           
            var premise = PremisesTxt.Text;
            var unit = unitNoTxt.Text;
            var lot = lotNoTxt.Text;
            var fromUnit = fromUnitNoTxt.Text;
            var toUnit = toUnitNoTxt.Text;
            if (string.IsNullOrEmpty(premise) || string.IsNullOrEmpty(lot))
                return false;
            else if (string.IsNullOrEmpty(unit) && (string.IsNullOrEmpty(fromUnit) || string.IsNullOrEmpty(toUnit)))
                return false;
            else
                return true;
        }

        private async void RemittanceSearchTask(string custName, string premise, string unit,string fromUnit,string toUnit, string lot)
        {
            progressbar1.Visibility = Visibility.Visible;
            var remittanceList= await Task.Run(() => {
                return svc.GetTdsRemitance(custName, premise, unit,fromUnit,toUnit, lot);
            });
            remittanceList = remittanceList.OrderBy(x => x.UnitNo).ToList();
            foreach (var x in remittanceList){  x.Show26qb = !x.IsDebitAdvice; };
            remitanceGrid.ItemsSource = remittanceList;
            TotalRecordsLbl.Content = remittanceList.Count;
            var totalTds = remittanceList.Sum(x => x.TdsAmount);
            TotalTDSLbl.Content = totalTds;
            progressbar1.Visibility = Visibility.Hidden;
          
        }
        
        private void TracesSearchFilter()
        {
            var remiitanceStatusID =( tracesRemitanceStatusddl.SelectedValue == null || Convert.ToInt32( tracesRemitanceStatusddl.SelectedValue) == -1) ? null : tracesRemitanceStatusddl.SelectedValue.ToString();
            var custName = tracesCustomerNameTxt.Text;
            var premise = tracesPremisesTxt.Text;
            var unit = tracesUnitNoTxt.Text;
            var lot = tracesLotNoTxt.Text;
            var fromUnit = tracesFromUnitNoTxt.Text;
            var toUnit = tracesToUnitNoTxt.Text;
            TracesSearchTask(custName, premise, unit,fromUnit,toUnit, lot, remiitanceStatusID);
        }

        private async void TracesSearchTask(string custName, string premise, string unit, string fromUnit, string toUnit, string lot,string remiittanceStatusID)
        {
            TracesProgressbar.Visibility = Visibility.Visible;
            var remittanceList = await Task.Run(() => {
                return svc.GetTdsPaidList(custName, premise, unit,fromUnit,toUnit, lot, remiittanceStatusID);
            });

            remittanceList = remittanceList.OrderBy(x => x.UnitNo).ToList();
            tdsRemitanceList =new ObservableCollection<TdsRemittanceDto>(remittanceList);
           
            //TracesGrid.ItemsSource = remittanceList;
            TracesGrid.ItemsSource = tdsRemitanceList;
            TracesProgressbar.Visibility = Visibility.Hidden;
        }

        private void Tracessearch() {
           // var remittanceStatusID = (tracesRemitanceStatusddl.SelectedValue == null || Convert.ToInt32(tracesRemitanceStatusddl.SelectedValue) == -1) ? null : tracesRemitanceStatusddl.SelectedValue.ToString();
            //var custName = tracesCustomerNameTxt.Text;
            //var premise = tracesPremisesTxt.Text;
            //var unit = tracesUnitNoTxt.Text;
            //var lot = tracesLotNoTxt.Text;
            //var fromUnit = tracesFromUnitNoTxt.Text;
            //var toUnit = tracesToUnitNoTxt.Text;

            var remittanceList = svc.GetTdsPaidList(custName, premise, unit, fromUnit, toUnit, lot, remittanceStatusID);
            remittanceList = remittanceList.OrderBy(x => x.UnitNo).ToList();
            tdsRemitanceList = new ObservableCollection<TdsRemittanceDto>(remittanceList);
            // Reload filter
            this.Dispatcher.Invoke((Action)(() =>
            {
                //TracesGrid.ItemsSource = remittanceList;
                TracesGrid.ItemsSource = tdsRemitanceList;

            }));
           
        }

        private void TracesFilter() {
            remittanceStatusID = (tracesRemitanceStatusddl.SelectedValue == null || Convert.ToInt32(tracesRemitanceStatusddl.SelectedValue) == -1) ? null : tracesRemitanceStatusddl.SelectedValue.ToString();
            custName = tracesCustomerNameTxt.Text;
             premise = tracesPremisesTxt.Text;
             unit = tracesUnitNoTxt.Text;
             lot = tracesLotNoTxt.Text;
             fromUnit = tracesFromUnitNoTxt.Text;
             toUnit = tracesToUnitNoTxt.Text;
        }

        private async void RequestForm16B(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var tdsremittanceModel = svc.GetTdsRemitanceById(model.ClientPaymentTransactionID);
            var reqNo = "";
            TracesFilter();
            TracesProgressbar.Visibility = Visibility.Visible;
            if (tdsremittanceModel != null) {
                await Task.Run(() => {
                    reqNo = FillTraces.AutoFillForm16B(tdsremittanceModel);
                });
            }
            //reqNo = FillTraces.AutoFillForm16B(tdsremittanceModel);
            TracesProgressbar.Visibility = Visibility.Hidden;
            if (reqNo != "")
            {
                var challanAmount = model.TdsAmount + model.TdsInterest + model.LateFee;
                Traces traces = new Traces(model, reqNo);
                traces.Owner = this;
                traces.ShowDialog();
            }
            else {
                MessageBox.Show("Request form16B Failed");
            }

            TracesSearchFilter();
        }
      
        private async void DownLoadForm(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var tdsremittanceModel = svc.GetTdsRemitanceById(model.ClientPaymentTransactionID);
            var remittanceModel = svc.GetRemitanceByTransID(model.ClientPaymentTransactionID);
            TracesFilter();
            if (tdsremittanceModel != null)
            {
                TracesProgressbar.Visibility = Visibility.Visible;
                await Task.Run(() => {
                    var filePath = FillTraces.AutoFillDownload(tdsremittanceModel, remittanceModel.F16BRequestNo, remittanceModel.DateOfBirth);

                    // Upload downloaded form
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        AutoUploadService autoUploadService = new AutoUploadService();
                        var status=autoUploadService.UploadForm16b(filePath, tdsremittanceModel);
                        if(!status)
                            MessageBox.Show("upload form Failed");
                    }
                    else {
                        MessageBox.Show("Download form Failed");
                    }
                    Tracessearch();
                });

                // FillTraces.AutoFillDownload(tdsremittanceModel, remittanceModel.F16BRequestNo, remittanceModel.DateOfBirth);
                TracesProgressbar.Visibility = Visibility.Hidden;
            }
        }

        private void UpdateRemittance(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var challanAmount = model.TdsAmount + model.TdsInterest + model.LateFee;

            Traces traces = new Traces(model);
            traces.Owner = this;
            traces.ShowDialog();
        }

        private async void DeleteFromRemittance(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var remittanceModel = svc.GetRemitanceByTransID(model.ClientPaymentTransactionID);
            if (remittanceModel.RemittanceID == 0) {
                MessageBox.Show("Remittance record is not yet created");
                return;
            }
            var resultMsg= MessageBox.Show("Are you sure to delete this?", "alert", MessageBoxButton.OKCancel);
            if (resultMsg == MessageBoxResult.OK) {
                progressbar1.Visibility = Visibility.Visible;
                bool status = false;
                await Task.Run(() => {
                    status = svc.DeleteRemittance(remittanceModel.RemittanceID);
                });
                progressbar1.Visibility = Visibility.Hidden;
                if (status)
                {
                    MessageBox.Show("Remittance is deleted successfully");
                    RemittanceSearchFilter();
                }
                else
                {
                    MessageBox.Show("Remittance is not deleted.");
                }
                
            }
        }

        private async void DeleteFromTrace(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var remittanceModel = svc.GetRemitanceByTransID(model.ClientPaymentTransactionID);
            if (remittanceModel.RemittanceID == 0)
            {
                MessageBox.Show("Remittance record is not yet created");
                return;
            }
            var resultMsg = MessageBox.Show("Are you sure to delete this?", "alert", MessageBoxButton.OKCancel);
            if (resultMsg == MessageBoxResult.OK)
            {
                TracesProgressbar.Visibility = Visibility.Visible;
                bool status = false ;
                await Task.Run(() => {
                    status= svc.DeleteRemittance(remittanceModel.RemittanceID);
                });
                TracesProgressbar.Visibility = Visibility.Hidden;
                if (status)
                {
                    MessageBox.Show("Remittance is deleted successfully");
                    TracesSearchFilter();
                }
                else
                {
                    MessageBox.Show("Remittance is not deleted.");
                }
            }
        }

        private async void SendMail(object sender, RoutedEventArgs e)
        {
            var model = (sender as Button).DataContext as TdsRemittanceDto;
            var challanAmount = model.TdsAmount + model.TdsInterest + model.LateFee;
            bool status = false;
            TracesProgressbar.Visibility = Visibility.Visible;
            await Task.Run(() => {
                status = svc.SendMail(model.ClientPaymentTransactionID);
            });
            TracesProgressbar.Visibility = Visibility.Hidden;
            if (status)
            {
                MessageBox.Show("Mail is delivered");
                TracesSearchFilter();
            }
            else
            {
                MessageBox.Show("Failed to send mail");
            }
        }

        private void accountddl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var combo = (ComboBox)(sender);
            if (combo.SelectedValue != null) {
                selectedAccount = Convert.ToInt32(combo.SelectedValue);

                var acct = accountList.Where(x => x.AccountId == Convert.ToInt32(accountddl.SelectedValue)).FirstOrDefault();
                selectedBank = acct.BankName;
                bankLogin = acct;
            }
        }

        private void fromUnitNoTxt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (unitNoTxt.Text != "") {
                unitNoTxt.Text = "";
            }

        }

        private void unitNoTxt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (fromUnitNoTxt.Text != "")
            {
                fromUnitNoTxt.Text = "";
                toUnitNoTxt.Text = "";
            }

        }

        private void tracesUnitNoTxt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (tracesFromUnitNoTxt.Text != "")
            {
                tracesFromUnitNoTxt.Text = "";
                tracesToUnitNoTxt.Text = "";
            }
        }

        private void tracesFromUnitNoTxt_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (tracesUnitNoTxt.Text != "")
            {
                tracesUnitNoTxt.Text = "";
            }

        }

        private void SelectAll_Checked(object sender, RoutedEventArgs e)
        {
            var isChecked = (sender as CheckBox).IsChecked;
            var isSelect=Convert.ToBoolean(isChecked);
            foreach (var item in tdsRemitanceList)
            {
               if(item.OnlyTDS)
                item.IsSelected = isSelect;
            }
            TracesGrid.ItemsSource = tdsRemitanceList;
            TracesGrid.Items.Refresh();
        }

        private void RemittanceAll_Checked(object sender, RoutedEventArgs e)
        {
            var isChecked = (sender as CheckBox).IsChecked;
            var isSelect = Convert.ToBoolean(isChecked);
            var remittanceList = (IList<TdsRemittanceDto>)remitanceGrid.ItemsSource;
            foreach (var item in remittanceList)
            {
                    item.IsSelected = isSelect;
            }
            remitanceGrid.ItemsSource = remittanceList;
            remitanceGrid.Items.Refresh();
        }

        private async void BulkRequest_Click(object sender, RoutedEventArgs e)
        {
            if (tdsRemitanceList == null)
                return;

            var selectedItems = tdsRemitanceList.ToList().FindAll(x => x.IsSelected == true);
            if (selectedItems.Count == 0)
                return;
            foreach (var item in selectedItems)
            {
              await proceedRequestAll(item);
            }

            Tracessearch();
            MessageBox.Show("Batch process completed.");
        }

        private async void BulkDownload_Click(object sender, RoutedEventArgs e)
        {
            if (tdsRemitanceList == null)
                return;

            var selectedItems = tdsRemitanceList.ToList().FindAll(x => x.IsSelected == true);
            if (selectedItems.Count == 0)
                return;
            foreach (var item in selectedItems)
            {
              await  proceedDownloadAll(item);
            }

            Tracessearch();
            MessageBox.Show("Batch process completed.");
        }

        private async Task<bool>  proceedRequestAll(TdsRemittanceDto model) {
            var tdsremittanceModel = svc.GetTdsRemitanceById(model.ClientPaymentTransactionID);
            var reqNo = "";
            TracesFilter();
            TracesProgressbar.Visibility = Visibility.Visible;
            if (tdsremittanceModel != null)
            {
                await Task.Run(() => {
                    reqNo = FillTraces.AutoFillForm16B(tdsremittanceModel);
                    if (reqNo != "")
                    {
                        AutoUploadService autoUploadService = new AutoUploadService();
                        autoUploadService.UpdateForm16BRequestNo(model.ClientPaymentTransactionID, reqNo);
                    }
                });
            }
            TracesProgressbar.Visibility = Visibility.Hidden;
           

            return true;
        }

        private async Task<bool> proceedDownloadAll(TdsRemittanceDto model)
        {
            var tdsremittanceModel = svc.GetTdsRemitanceById(model.ClientPaymentTransactionID);
            var remittanceModel = svc.GetRemitanceByTransID(model.ClientPaymentTransactionID);
            TracesFilter();
            if (tdsremittanceModel != null)
            {
                TracesProgressbar.Visibility = Visibility.Visible;
                await Task.Run(() => {
                    var filePath = FillTraces.AutoFillDownload(tdsremittanceModel, remittanceModel.F16BRequestNo, remittanceModel.DateOfBirth);

                    // Upload downloaded form
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        AutoUploadService autoUploadService = new AutoUploadService();
                        autoUploadService.UploadForm16b(filePath, tdsremittanceModel);
                    }
                   
                });

                TracesProgressbar.Visibility = Visibility.Hidden;
                
            }
            return true;
        }
        private void bankddl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var combo = (ComboBox)(sender);
            if (combo.SelectedValue != null)
            {
                selectedBank = combo.SelectedValue.ToString();
            }
        }

        private void ResetOtp_Click(object sender, RoutedEventArgs e)
        {
            if (bankLogin == null)
            {
                MessageBox.Show("Please select account.");
                return;
            }

            if (bankLogin.LaneNo == null)
            {
                MessageBox.Show("Lane No is not available for this account");
                return;
            }

            var status = svc.DeleteOTP(bankLogin.LaneNo.Value);
            if (status)
                MessageBox.Show("OTP reset is done");
            else
                MessageBox.Show("OTP reset is failed");
        }

        private void challan_download_Click(object sender, RoutedEventArgs e)
        {
            var remittanceList = (List<TdsRemittanceDto>)remitanceGrid.ItemsSource;
            remittanceList = remittanceList.Where(x => x.IsSelected == true).ToList();
            var id = remittanceList[0].ClientPaymentTransactionID;
            AutoFillDto autoFillDto = svc.GetAutoFillData(id);
            FillForm26QB_ICICI.DownloadChallanFromTaxPortal(autoFillDto,id);
        }
    }
}
