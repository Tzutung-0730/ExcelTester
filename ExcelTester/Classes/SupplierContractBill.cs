using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelTester.Attributes;
using OfficeOpenXml;

namespace ExcelTester.Classes
{
    internal class SupplierContractBill
    {
        public static void WriteExcel_Invoice(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnSupplierContractBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];

            var bill = excelData.First();

            // 結算日期
            sheet1.Cells[4, 14].Value = $"{bill.SettleTime:yyy/MM/dd}";
            
            // 事由
            sheet1.Cells[5, 3].Value = $"{bill.BillYear - 1911}/{bill.BillMonth:D2} 份再生能源電能費用({bill.SupplierName})";

            // 專案工作代號，以「、」區隔
            var projectCodes = string.Join("、", bill.Items
                .Where(i => !string.IsNullOrEmpty(i.ProjectId)) // 過濾掉空值
                .Select(i => i.ProjectId)
                .Distinct()); // 過濾重複值
            sheet1.Cells[6, 10].Value = $"專案工作代號：{projectCodes}";

            // 說明第1點
            var supplierCodes = string.Join("、", bill.Items.Select(i => i.SupplierPlaceId));
            sheet1.Cells[9, 2].Value = $"1、{bill.SupplierName}案場電號{supplierCodes}，再生能源電能費用計 NT${bill.Items.Sum(i => i.TotalAmount):N0}元(含稅)。";

            // 說明第2點
            sheet1.Cells[11, 2].Value = $"  戶名：{bill.ReceiveName}";
            sheet1.Cells[12, 2].Value = $"  行庫：{bill.ReceiveBankName}";
            sheet1.Cells[13, 2].Value = $"  匯款帳號：{bill.ReceiveAccount}";

            // 說明第3點
            sheet1.Cells[14, 2].Value = $"3、附件：驗收紀錄表、台汽電綠能付款通知單、發票、{bill.SupplierName}存摺影本。";

            // 填寫付款資訊到 sheet1
            sheet1.Cells[19, 2].Value = bill.SupplierName;
            sheet1.Cells[19, 6].Value = bill.PaymentMethod;
            sheet1.Cells[19, 8].Value = bill.Items.Sum(i => i.TotalAmount);

            package.SaveAs(new FileInfo(newFileName));
        }

        public static void WriteExcel_PaymentNotice(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnSupplierContractBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];

            var bill = excelData.First();

            sheet1.Cells[5, 4].Value = bill.SupplierContractBillId;
            sheet1.Cells[5, 10].Value = $"製發日期：{bill.SettleTime:yyyy/MM/dd}";
            sheet1.Cells[5, 10, 5, 11].Merge = false;

            var startDate = new DateTime(bill.BillYear, bill.BillMonth, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            sheet1.Cells[5, 7].Value = $"{startDate:yyyy/MM/dd}~{endDate:yyyy/MM/dd}";

            sheet1.Cells[6, 4].Value = bill.SupplierName;
            sheet1.Cells[7, 4].Value = bill.SupplierTaxIDNumber;
            sheet1.Cells[8, 4].Value = bill.SupplierAddress;
            sheet1.Cells[9, 4].Value = bill.SupplierContactNumber;
            sheet1.Cells[9, 10].Value = $"付款期限：{bill.PaymentDeadline:yyyy/MM/dd}";

            int startRow = 12;
            int currentRow = 12;
            int maxRowsPerPage = 6;

            foreach (var Item in bill.Items)
            {
                if (currentRow >= startRow + maxRowsPerPage)
                {
                    sheet1.InsertRow(currentRow, 1); // 插入一行
                    sheet1.Cells[currentRow, 3, currentRow, 4].Merge = true;
                    sheet1.Cells[currentRow, 10, currentRow, 12].Merge = true;

                    for (int col = 1; col <= 12; col++)
                    {
                        sheet1.Cells[currentRow, col].StyleID = sheet1.Cells[currentRow - 1, col].StyleID;
                    }
                }

                sheet1.Cells[currentRow, 8, currentRow, 9].Merge = true;
                sheet1.Cells[currentRow, 2].Value = currentRow - startRow + 1;
                sheet1.Cells[currentRow, 3].Value = Item.ItemName;
                sheet1.Cells[currentRow, 5].Value = Item.Unit;
                sheet1.Cells[currentRow, 6].Value = Item.UnitPrice;
                sheet1.Cells[currentRow, 7].Value = Item.Quantity;
                sheet1.Cells[currentRow, 8].Value = Item.TotalAmount;
                sheet1.Cells[currentRow, 10].Value = Item.Note;

                currentRow++;
            }

            if (currentRow < 18) currentRow = 18;
 
            sheet1.Cells[currentRow, 8].Formula = $"=SUM(H{startRow}:H{currentRow - 1})";
            sheet1.Cells[currentRow + 1, 8].Formula = $"=ROUND(H{currentRow}*5%,0)";
            sheet1.Cells[currentRow + 2, 8].Formula = $"=H{currentRow} + H{currentRow + 1}";
            sheet1.Cells[currentRow + 7, 1].Value = $"付款方式 ： {bill.PaymentMethod}";
            sheet1.Cells[currentRow + 8, 1].Value = $"收款戶名 ： {bill.ReceiveName}";
            sheet1.Cells[currentRow + 9, 1].Value = $"收款行庫 ： {bill.ReceiveBankName}";
            sheet1.Cells[currentRow + 10, 1].Value = $"收款帳號 ： {bill.ReceiveAccount}";

            package.SaveAs(new FileInfo(newFileName));
        }
    }
}
