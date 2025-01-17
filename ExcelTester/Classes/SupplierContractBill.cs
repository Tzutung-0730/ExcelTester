using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelTester.Attributes;
using OfficeOpenXml;

namespace ExcelTester.Classes
{
    internal class SupplierContractBill
    {
        public static void WriteExcel1(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnSupplierContractBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];

            var bill = excelData.First();

            // 結算日期
            sheet1.Cells[4, 14].Value = ConvertToTaiwanCalendar(bill.SettleTime, "yyy/MM/dd");
            
            // 事由
            sheet1.Cells[5, 3].Value = $"{ConvertToTaiwanCalendar(new DateTime(bill.BillYear, bill.BillMonth, 1), "yyy/MM")} 份再生能源電能費用({bill.SupplierName})";
            
            // 專案工作代號，以「、」區隔
            var projectCodes = string.Join("、", bill.Items.Select(i => i.SupplierContractBillId));
            sheet1.Cells[6, 11].Value = $"專案工作代號：{projectCodes}";

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
            sheet1.Cells[19, 6].Value = "轉帳";
            sheet1.Cells[19, 8].Value = bill.Items.Sum(i => i.TotalAmount).ToString("#,##0");

            package.SaveAs(new FileInfo(newFileName));
        }

        public static void WriteExcel2(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnSupplierContractBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];

            var bill = excelData.First();

            sheet1.Cells[5, 4].Value = bill.SupplierContractBillId;
            sheet1.Cells[5, 12].Value = "yyyy/MM/dd";

            var startDate = new DateTime(bill.BillYear, bill.BillMonth, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            sheet1.Cells[5, 7].Value = $"{startDate:yyyy/MM/dd}~{endDate:yyyy/MM/dd}";

            sheet1.Cells[6, 4].Value = bill.SupplierName;
            sheet1.Cells[7, 4].Value = "統一編號";
            sheet1.Cells[8, 4].Value = "地址";
            sheet1.Cells[9, 4].Value = "電話";
            sheet1.Cells[9, 10].Value = $"付款期限：{bill.PaymentDeadline:yyyy/MM/dd}";

            int startRow = 12;
            int currentRow = 12;

            foreach (var Item in bill.Items)
            {
                sheet1.Cells[currentRow, 2].Value = currentRow - startRow + 1;
                sheet1.Cells[currentRow, 3].Value = Item.ItemName;
                sheet1.Cells[currentRow, 5].Value = Item.Unit;
                sheet1.Cells[currentRow, 6].Value = Item.UnitPrice;
                sheet1.Cells[currentRow, 7].Value = Item.Quantity;
                sheet1.Cells[currentRow, 8].Value = Item.TotalAmount;
                sheet1.Cells[currentRow, 9].Value = Item.Note;

                for (int col = 1; col <= 5; col++)
                {
                    var cell = sheet1.Cells[currentRow, col];

                    if (col == 6 || col == 8)
                    {
                        cell.Style.Numberformat.Format = "#,##0"; // 千分位格式
                    }
                }

                currentRow++;

                if (currentRow - startRow + 1 > 5)
                {
                    sheet1.InsertRow(currentRow + 1, 1);
                }
            }

            if (currentRow > 16) 
            {
                sheet1.DeleteRow(currentRow + 1);
            } else
            {
                currentRow = 17;
            }
 
            sheet1.Cells[18, 8].Formula = $"SUM(H12:H{currentRow})";
            sheet1.Cells[19, 8].Formula = $"=H{currentRow + 1} * 0.05";
            sheet1.Cells[20, 8].Formula = $"=H{currentRow + 1} + H{currentRow + 2}";

            package.SaveAs(new FileInfo(newFileName));
        }

        private static string ConvertToTaiwanCalendar(DateTime date, string format)
        {
            var taiwanCalendar = new System.Globalization.TaiwanCalendar();
            return date.ToString(format, new System.Globalization.CultureInfo("zh-TW")
            {
                DateTimeFormat = { Calendar = taiwanCalendar }
            });
        }
    }
}
