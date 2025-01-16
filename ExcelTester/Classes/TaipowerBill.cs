using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelTester.Attributes;
using OfficeOpenXml;

namespace ExcelTester.Classes
{
    public class TaipowerBill
    {
        public static void GenerateTaipowerBill(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnTaipowerBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));


            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];
            int paymentRowStart = 18;

            // 填寫結算日期到 N4，轉為民國格式
            foreach (var bill in excelData)
            {
                // 結算日期：轉為民國年格式
                sheet1.Cells["N4"].Value = ConvertToTaiwanCalendar(bill.SettleTime, "yyy/MM/dd");
                sheet1.Cells["N4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                // 事由欄位
                sheet1.Cells["C5"].Value = $"繳納 {ConvertToTaiwanCalendar(new DateTime(bill.BillYear, bill.BillMonth, 1), "yyy/MM")} 轉供費用";

                // 說明第 1 點
                sheet1.Cells["B9"].Value = $"1、繳納 {ConvertToTaiwanCalendar(new DateTime(bill.BillYear, bill.BillMonth, 1), "yyy/MM")} 台電轉供費用，總計 NT${bill.Items[0].TotalAmount:N0} (含稅)，詳見清單。\n";

                // 說明第 2 點
                //sheet1.Cells["E10"].Value = "2、匯款資料\n\t戶名：台灣電力股份有限公司\n\t行庫：台灣銀行公館分行\n\t帳號：詳見公用槽之台電轉供輸電費繳費單";
                //sheet1.Cells["E10"].Style.WrapText = true; // 啟用文字換行

                // 說明第 3 點
                sheet1.Cells["B14"].Value = $"3、請財務部安排於 {ConvertToTaiwanCalendar(bill.PaymentDeadline, "yyy/MM/dd")} 前繳納";

                foreach (var item in bill.Items)
                {
                    sheet1.InsertRow(paymentRowStart, 1);

                    // 合併 B-E 欄，填入付款對象 (SupplierPlaceName)
                    sheet1.Cells[paymentRowStart, 2, paymentRowStart, 5].Merge = true; // 合併 B-E 欄
                    sheet1.Cells[paymentRowStart, 2].Value = item.SupplierPlaceName; // 填入供應商名稱
                    // 填入 F 欄 (付款方式)    
                    sheet1.Cells[paymentRowStart, 6].Value = "轉帳"; // 固定為 "轉帳"
                    sheet1.Cells[paymentRowStart, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 水平置中
                    sheet1.Cells[paymentRowStart, 7].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center; // 垂直置中
                    sheet1.Cells[paymentRowStart, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); // 邊框
                    // 填入 G 欄 (幣別)
                    sheet1.Cells[paymentRowStart, 7].Value = "NTD"; // 固定為 "NTD"
                    sheet1.Cells[paymentRowStart, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 水平置中
                    sheet1.Cells[paymentRowStart, 7].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center; // 垂直置中
                    sheet1.Cells[paymentRowStart, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); // 邊框
                    // 填入 H 欄 (付款金額)
                    sheet1.Cells[paymentRowStart, 8].Value = item.TotalAmount; // 使用 Item 的 TotalAmount
                    sheet1.Cells[paymentRowStart, 8].Style.Numberformat.Format = "#,##0"; // 千分位格式
                    sheet1.Cells[paymentRowStart, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center; // 水平置中
                    sheet1.Cells[paymentRowStart, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center; // 垂直置中
                    sheet1.Cells[paymentRowStart, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); // 邊框
                    paymentRowStart++;
                }
            }


            ExcelWorksheet sheet2 = package.Workbook.Worksheets[1];
            int currentRow = 2;

            foreach (var item in excelData)
            {
                ItemToRow(sheet2, currentRow, item);
                currentRow++;
            }

            // 將「總計金額」行插入到最後一行
            sheet2.Cells[currentRow, 1].Value = "總計金額";
            sheet2.Cells[currentRow, 4].Formula = $"SUM(D2:D{currentRow - 1})"; // 設定總計公式
            sheet2.Cells[currentRow, 5].Formula = $"SUM(E2:E{currentRow - 1})"; // 設定總計公式

            // 設定總計行的樣式
            sheet2.Cells[currentRow, 1, currentRow, 5].Style.Font.Bold = true;
            sheet2.Cells[2, 4, currentRow, 5].Style.Numberformat.Format = "#,##0";

            // 儲存 Excel 到指定的新檔案名稱
            package.SaveAs(new FileInfo(newFileName));
        }

        private static void ItemToRow(ExcelWorksheet sheet, int row, ExcelColumnTaipowerBill item)
        {
            // 獲取物件的所有屬性，並按屬性順序填入對應的 Excel 列
            var properties = item.GetType().GetProperties();

            for (int col = 1; col <= properties.Length; col++)
            {
                var prop = properties[col - 1]; // 根據屬性順序決定對應的列
                var value = prop.GetValue(item);

                sheet.Cells[row, col].Value = value;

                // 如果是數字欄位，統一設置數字格式
                if (value is int || value is decimal || value is double)
                {
                    sheet.Cells[row, col].Style.Numberformat.Format = "#,##0"; // 千分位格式
                }

                // 設定資料行樣式
                if (row > 2) 
                {
                    sheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    sheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#F2F2F2"));
                    sheet.Cells[row, col].Style.Font.Bold = false;
                    sheet.Cells[row, col].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium); // 加框線
                }
            }
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
