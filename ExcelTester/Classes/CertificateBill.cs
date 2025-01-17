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
    public class CertificateBill
    {
        public static void WriteExcel(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnCertificateBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];
            ExcelWorksheet sheet2 = package.Workbook.Worksheets[1];

            int sheet2CurrentRow = 2;
            var bill = excelData.First();

            // 結算日期
            sheet1.Cells[4, 14].Value = ConvertToTaiwanCalendar(bill.SettleTime, "yyy/MM/dd");

            // 事由
            sheet1.Cells[5, 3].Value = $"{ConvertToTaiwanCalendar(new DateTime(bill.BillYear, bill.BillMonth, 1), "yyy/MM")} 份 再生能源憑證規費";

            // 說明第1點
            sheet1.Cells[9, 2].Value = $"1、繳納 {ConvertToTaiwanCalendar(new DateTime(bill.BillYear, bill.BillMonth, 1), "yyy/MM")} 再生能源憑證規費，審審查費總計 NTD${bill.Items.Sum(i => i.CertificateFee):N0}，服務費總計 NTD${bill.Items.Sum(i => i.ServiceFee):N0}，共計 NTD${bill.Items.Sum(i => i.CertificateFee + i.ServiceFee):N0}，詳見清單。";

            // 說明第3點
            sheet1.Cells[14, 2].Value = $"3、請財務部安排於 {ConvertToTaiwanCalendar(bill.PaymentDeadline, "yyy/MM/dd")} 前繳納（憑證中心同意）。";

            foreach (var item in bill.Items)
            {
                // 將資料填入 sheet2
                sheet2.Cells[sheet2CurrentRow, 1].Value = item.ProjectId;
                sheet2.Cells[sheet2CurrentRow, 2].Value = item.SupplierPlaceName;
                sheet2.Cells[sheet2CurrentRow, 3].Value = item.CertificateCount;
                sheet2.Cells[sheet2CurrentRow, 4].Value = item.CertificateFee;
                sheet2.Cells[sheet2CurrentRow, 5].Value = item.ServiceFee;

                for (int col = 1; col <= 5; col++)
                {
                    var cell = sheet2.Cells[sheet2CurrentRow, col];

                    // 設定資料行樣式
                    if (sheet2CurrentRow > 2)
                    {
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#F2F2F2"));
                        cell.Style.Font.Bold = false;
                        cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    }

                    // 設定數字格式
                    if (col == 4 || col == 5)
                    {
                        cell.Style.Numberformat.Format = "#,##0"; // 千分位格式
                    }
                }

                sheet2.Cells[sheet2CurrentRow, 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                sheet2.Cells[sheet2CurrentRow, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                sheet2CurrentRow++;
            }

            // 設定 2A 到 2E 單元格的下框線為 Thin
            sheet2.Cells[2, 1, 2, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            sheet2.Cells[sheet2CurrentRow - 1, 1, sheet2CurrentRow - 1, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // 填寫總計行到 sheet2
            sheet2.Cells[sheet2CurrentRow, 1].Value = "總計金額";
            sheet2.Cells[sheet2CurrentRow, 4].Formula = $"SUM(D2:D{sheet2CurrentRow - 1})";
            sheet2.Cells[sheet2CurrentRow, 5].Formula = $"SUM(E2:E{sheet2CurrentRow - 1})";

            // 填寫付款金額到 sheet1
            sheet1.Cells[17, 8].Formula = $"=清單!D{sheet2CurrentRow} + 清單!E{sheet2CurrentRow}";

            // 設定總計行的樣式
            sheet2.Cells[sheet2CurrentRow, 1, sheet2CurrentRow, 5].Style.Font.Bold = true;
            sheet2.Cells[2, 4, sheet2CurrentRow, 5].Style.Numberformat.Format = "#,##0";

            // 儲存 Excel 到指定的新檔案名稱
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
