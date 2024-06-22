using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SpreadsheetLight;

namespace AutoFill
{
   public class ExportToExcel
   {
      
        public void RemittanceExport(List<TdsRemittanceDto> remList)
        {
            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            var filePath = @downloadPath + "\\TDS-Remittance"+DateTime.Now.ToString("-ddMMyyhhmmss")+ ".xlsx";
            SLDocument sl = new SLDocument();
            //headers
            sl.SetCellValue(1, 1, "Unit No");
            sl.SetCellValue(1, 2, "Lot No");
            sl.SetCellValue(1, 3, "Customer Name");
            sl.SetCellValue(1, 4, "Seller Name");
            sl.SetCellValue(1, 5, "Property Name");
            sl.SetCellValue(1, 6, "Amount");
            sl.SetCellValue(1, 7, "TDS");
            sl.SetCellValue(1, 8, "TDS Interest");
            sl.SetCellValue(1, 9, "Late Fee");
            sl.SetCellValue(1, 10, "Gross Amount");
            sl.SetCellValue(1, 11, "Deduction Date");
            sl.SetCellValue(1, 12, "Is DA Completed");
            sl.SetCellValue(1, 13, "Transaction ID");
            sl.SetCellValue(1, 14, "Remark");
            sl.SetCellValue(1, 15, "CIN No");
            sl.SetCellValue(1, 16, "Transaction Log");

            var length = remList.Count;
            for (var i=0; i< length; i++)
            {
                var obj = remList[i];
                sl.SetCellValue(i+2, 1, obj.UnitNo);
                sl.SetCellValue(i+2, 2, obj.LotNo);
                sl.SetCellValue(i+2, 3, obj.CustomerName);
                sl.SetCellValue(i+2, 4, obj.SellerName);
                sl.SetCellValue(i + 2, 5, obj.PropertyPremises);
                sl.SetCellValue(i+2, 6, obj.AmountShare);
                sl.SetCellValue(i+2, 7, obj.TdsAmount);
                sl.SetCellValue(i+2, 8, obj.TdsInterest);
                sl.SetCellValue(i+2, 9, obj.LateFee);
                sl.SetCellValue(i+2, 10, obj.GrossAmount);
                sl.SetCellValue(i+2, 11, obj.DateOfDeduction.ToString("dd-MMM-yyyy"));
                sl.SetCellValue(i+2, 12, obj.IsDebitAdvice);
                sl.SetCellValue(i+2, 13, obj.ClientPaymentTransactionID);
                sl.SetCellValue(i + 2, 14, obj.RemarkDesc);
                sl.SetCellValue(i + 2, 15, obj.CinNo);
                sl.SetCellValue(i + 2, 16, obj.TransactionLog);
            }
         
            sl.SaveAs(filePath);

        }

        public void TracesExport(List<TracesModel> remList)
        {
            var downloadPath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            var filePath = @downloadPath + "\\TDS-Traces" + DateTime.Now.ToString("-ddMMyyhhmmss") + ".xlsx";
            SLDocument sl = new SLDocument();
            //headers
         
            sl.SetCellValue(1, 1, "Customer Name");
            sl.SetCellValue(1, 2, "Seller Name");
            sl.SetCellValue(1, 3, "Property Name");
            sl.SetCellValue(1, 4, "Unit No");
            sl.SetCellValue(1, 5, "Lot No");
            sl.SetCellValue(1, 6, "Challan Amount");
            sl.SetCellValue(1, 7, "Acknowledgement No");
            sl.SetCellValue(1, 8, "Challan Date");
            sl.SetCellValue(1, 9, "Request Date");
            sl.SetCellValue(1, 10, "Request No");
            sl.SetCellValue(1, 11, "status");
            sl.SetCellValue(1, 12, "Transaction ID");
            sl.SetCellValue(1, 13, "Remark");
          

            var length = remList.Count;
            for (var i = 0; i < length; i++)
            {
                var obj = remList[i];
                sl.SetCellValue(i + 2, 1, obj.CustomerName);
                sl.SetCellValue(i + 2, 2, obj.SellerName);
                sl.SetCellValue(i + 2, 3, obj.PropertyPremises);
                sl.SetCellValue(i + 2, 4, obj.UnitNo);
                sl.SetCellValue(i + 2, 5, obj.LotNo);
                sl.SetCellValue(i + 2, 6, obj.ChallanAmount);
                sl.SetCellValue(i + 2, 7, obj.ChallanAckNo);
                sl.SetCellValue(i + 2, 8, obj.ChallanDate.ToString("dd-MMM-yyyy"));
                sl.SetCellValue(i + 2, 9, (obj.F16BDateOfReq==null)?"":obj.F16BDateOfReq.Value.ToString("dd-MMM-yyyy"));
                sl.SetCellValue(i + 2, 10, obj.F16BRequestNo);
                sl.SetCellValue(i + 2, 11, obj.RemittanceStatus);
                sl.SetCellValue(i + 2, 12, obj.ClientPaymentTransactionID);
                sl.SetCellValue(i + 2, 13, obj.RemarkDesc);
            }

            sl.SaveAs(filePath);

        }
    }
}
