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
    public class TaipowerBill
    {
        public static void WriteExcel(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnTaipowerBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];
            ExcelWorksheet sheet2 = package.Workbook.Worksheets[1];

            int sheet2CurrentRow = 2;
            var bill = excelData.First();
            var totalWithoutTax = bill.Items.Sum(i => Math.Round(i.TotalAmount / 1.25, 0));
            var totalWithTax = bill.Items.Sum(i => i.TotalAmount);

            // 結算日期
            sheet1.Cells[4, 14].Value = $"{bill.SettleTime.Year - 1911}/{bill.SettleTime:MM/dd}";

            // 事由
            sheet1.Cells[5, 3].Value = $"繳納 {bill.BillYear - 1911}/{bill.BillMonth:D2} 轉供費用";

            // 說明第1點
            sheet1.Cells[9, 2].Value = $"1、繳納 {bill.BillYear - 1911}/{bill.BillMonth:D2} 台電轉供費用，總計 NT${totalWithTax:N0} (含稅)，詳見清單。";
                
            // 說明第3點
            sheet1.Cells[14, 2].Value = $"3、請財務部安排於 {bill.PaymentDeadline.Year - 1911}/{bill.PaymentDeadline:MM/dd} 前繳納。";

            foreach (var item in bill.Items)
            {
                // 將資料填入 sheet2
                sheet2.Cells[sheet2CurrentRow, 1].Value = item.ProjectId;
                sheet2.Cells[sheet2CurrentRow, 2].Value = item.SupplierPlaceName;
                sheet2.Cells[sheet2CurrentRow, 3].Value = item.CustomerName;
                sheet2.Cells[sheet2CurrentRow, 4].Value = Math.Round(item.TotalAmount / 1.25, 0); ;
                sheet2.Cells[sheet2CurrentRow, 5].Value = item.TotalAmount;

                for (int col = 1; col <= 5; col++)
                {
                    sheet2.Cells[sheet2CurrentRow, col].StyleID = sheet2.Cells[2, col].StyleID;
                    sheet2.Cells[sheet2CurrentRow, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                sheet2CurrentRow++;
            }

            sheet2.Cells[sheet2CurrentRow - 1, 1, sheet2CurrentRow - 1, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // 填寫總計行到 sheet2
            sheet2.Cells[sheet2CurrentRow, 1].Value = "總計金額";
            sheet2.Cells[sheet2CurrentRow, 4].Value = totalWithoutTax;
            sheet2.Cells[sheet2CurrentRow, 5].Value = totalWithTax;

            // 填寫付款金額到 sheet1
            sheet1.Cells[17, 8].Value = totalWithTax;

            // 設定總計行的樣式
            sheet2.Cells[sheet2CurrentRow, 1, sheet2CurrentRow, 5].Style.Font.Bold = true;
            sheet2.Cells[2, 4, sheet2CurrentRow, 5].Style.Numberformat.Format = "#,##0";

            // 儲存 Excel 到指定的新檔案名稱
            package.SaveAs(new FileInfo(newFileName));
        }
    }
}
