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
            sheet1.Cells[4, 14].Style.Font.Name = "標楷體";
            sheet1.Cells[4, 14].Style.Font.Size = 16;
            sheet1.Cells[4, 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            // 事由
            sheet1.Cells[5, 3].Value = $"{ConvertToTaiwanCalendar(new DateTime(bill.BillYear, bill.BillMonth, 1), "yyy/MM")} 份再生能源電能費用([電業公司])";
            sheet1.Cells[5, 3].Style.Font.Name = "標楷體";
            sheet1.Cells[5, 3].Style.Font.Size = 16;

            // 專案工作代號，以「、」區隔
            var projectCodes = string.Join("、", bill.Items.Select(i => i.SupplierContractBillId));
            sheet1.Cells[6, 11].Value = $"專案工作代號：{projectCodes}";
            sheet1.Cells[6, 11].Style.Font.Name = "標楷體";
            sheet1.Cells[6, 11].Style.Font.Size = 12;

            // 說明第1點
            var supplierCodes = string.Join("、", bill.Items.Select(i => i.SupplierPlaceId));
            sheet1.Cells[9, 2].Value = $"1、{bill.SupplierName}案場電號{supplierCodes}，再生能源電能費用計 NT${bill.Items.Sum(i => i.TotalAmount):N0}元(含稅)。";
            sheet1.Cells[9, 2].Style.Font.Name = "標楷體";
            sheet1.Cells[9, 2].Style.Font.Size = 16;

            // 說明第2點
            sheet1.Cells[11, 2].Value = $"  戶名：{bill.ReceiveName}";
            sheet1.Cells[11, 2].Style.Font.Name = "標楷體";
            sheet1.Cells[11, 2].Style.Font.Size = 16;
            sheet1.Cells[12, 2].Value = $"  行庫：{bill.ReceiveBankName}";
            sheet1.Cells[12, 2].Style.Font.Name = "標楷體";
            sheet1.Cells[12, 2].Style.Font.Size = 16;
            sheet1.Cells[13, 2].Value = $"  匯款帳號：{bill.ReceiveAccount}";
            sheet1.Cells[13, 2].Style.Font.Name = "標楷體";
            sheet1.Cells[13, 2].Style.Font.Size = 16;

            // 說明第3點
            sheet1.Cells[14, 2].Value = $"3、附件：驗收紀錄表、台汽電綠能付款通知單、發票、[電業公司]存摺影本。";
            sheet1.Cells[14, 2].Style.Font.Name = "標楷體";
            sheet1.Cells[14, 2].Style.Font.Size = 16;
        }

        public static void WriteExcel2(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnSupplierContractBill> excelData)
        {

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
