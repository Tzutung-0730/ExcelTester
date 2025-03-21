﻿using System;
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
using OfficeOpenXml.Style;

namespace ExcelTester.Classes
{
    internal class CustomerContractBill
    {
        public static void WriteExcel_Invoice(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnCustomerContractBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];

            var bill = excelData.First();
            int startRow = 8;
            int currentRow = 8;
            int maxRowsPerPage = 4;
            var startDate = new DateTime(bill.BillYear, bill.BillMonth, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            var totalAmount = bill.Items.Sum(i => i.TotalAmount);
            var tax = Math.Round(totalAmount * 0.05, 0);
            var totalWithTax = totalAmount + tax;

            sheet1.Cells[3, 10].Value = $"{bill.SettleTime.Year - 1911}/{bill.SettleTime:MM/dd}";

            // 專案代號，以「、」區隔
            var projectCodes = string.Join("、", bill.Items
                            .Where(i => !string.IsNullOrEmpty(i.ProjectId)) // 過濾掉空值
                            .Select(i => i.ProjectId)
                            .Distinct()); // 過濾重複值
            sheet1.Cells[4, 6].Value = $"專案代號：{projectCodes}";
            sheet1.Cells[5, 1].Value = $"買    受    人：{bill.CustomerName}";
            sheet1.Cells[5, 6].Value = $"統一編號：{bill.CustomerTaxIDNumber}";
            sheet1.Cells[6, 1].Value = $"地      　　址：{bill.CustomerAddress}";

            foreach (var Item in bill.Items)
            {
                if (currentRow >= startRow + maxRowsPerPage)
                {
                    sheet1.InsertRow(currentRow, 1); // 插入一行
                    sheet1.Cells[currentRow, 1, currentRow, 2].Merge = true;
                    sheet1.Cells[currentRow, 3, currentRow, 4].Merge = true;
                }

                for (int col = 1; col <= 8; col++)
                {
                    sheet1.Cells[currentRow, col].StyleID = sheet1.Cells[currentRow - 1, col].StyleID;
                }

                sheet1.Cells[currentRow, 5, currentRow, 6].Merge = true;
                sheet1.Cells[currentRow, 7, currentRow, 8].Merge = true;

                if (Item.ItemName == "再生能源發電費" || Item.ItemName == "電能代輸轉供服務費用")
                {
                    sheet1.Row(currentRow).Height = 39;
                    sheet1.Cells[currentRow, 1].Value = $"{Item.ItemName}\n({startDate.Year - 1911}/{startDate:MM/dd}~{endDate.Year - 1911}/{endDate:MM/dd})";
                    sheet1.Cells[currentRow, 1].Style.WrapText = true; // 啟用自動換行
                }
                else
                {
                    sheet1.Row(currentRow).Height = 25.2;
                    sheet1.Cells[currentRow, 1].Value = Item.ItemName;
                }

                sheet1.Cells[currentRow, 3].Value = Item.Quantity;
                sheet1.Cells[currentRow, 5].Value = Item.UnitPrice;
                sheet1.Cells[currentRow, 7].Value = Item.TotalAmount;
                sheet1.Cells[currentRow, 3].Style.Numberformat.Format = "#,##0";
                double value = Convert.ToDouble(sheet1.Cells[currentRow, 5].Value);
                if (value % 1 == 0) // 判斷是否為整數
                {
                    sheet1.Cells[currentRow, 5].Style.Numberformat.Format = "#,##0"; // 整數格式
                } else
                {
                    sheet1.Cells[currentRow, 5].Style.Numberformat.Format = "#,##0.####"; // 小數格式
                }
                sheet1.Cells[currentRow, 7].Style.Numberformat.Format = "#,##0";
                sheet1.Cells[currentRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet1.Cells[currentRow, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet1.Cells[currentRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                currentRow++;
            }
            
            if (currentRow < 12) currentRow = 12;

            sheet1.Cells[currentRow, 8].Value = totalAmount;
            sheet1.Cells[currentRow + 1, 8].Value = tax;
            sheet1.Cells[currentRow + 1, 8].Style.Numberformat.Format = "#,##0";
            sheet1.Cells[currentRow + 2, 8].Value = totalWithTax;
            sheet1.Cells[currentRow + 1, 10].Value = $"{bill.PaymentDeadline.Year - 1911}/{bill.PaymentDeadline:MM/dd}";
            sheet1.Cells[currentRow + 1, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


            package.SaveAs(new FileInfo(newFileName));
        }

        public static void WriteExcel_PaymentNotice(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnCustomerContractBill> excelData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(srcTemplateFile));

            ExcelWorksheet sheet1 = package.Workbook.Worksheets[0];

            var bill = excelData.First();
            var startDate = new DateTime(bill.BillYear, bill.BillMonth, 1);
            var endDate = startDate.AddMonths(1).AddDays(-1);
            var totalAmount = bill.Items.Sum(i => i.TotalAmount);
            var tax = Math.Round(totalAmount * 0.05, 0);
            var totalWithTax = totalAmount + tax;

            sheet1.Cells[4, 4].Value = bill.BillCode;
            sheet1.Cells[3, 9].Value = $"製發日期：{bill.SettleTime:yyyy/MM/dd}";
            sheet1.Cells[4, 7, 4, 8].Merge = false;            
            sheet1.Cells[4, 5].Value = $"計費期間：{startDate:yyyy/MM/dd}~{endDate:yyyy/MM/dd}";

            sheet1.Cells[5, 4].Value = bill.CustomerName;
            sheet1.Cells[6, 4].Value = bill.CustomerTaxIDNumber;
            sheet1.Cells[7, 4].Value = bill.CustomerAddress;
            sheet1.Cells[8, 4].Value = bill.CustomerContactNumber;
            sheet1.Cells[8, 8].Value = $"繳費期限：{bill.PaymentDeadline:yyyy/MM/dd}";

            int startRow = 11;
            int currentRow = 11;
            int maxRowsPerPage = 4;

            foreach (var Item in bill.Items)
            {
                if (currentRow >= startRow + maxRowsPerPage)
                {
                    sheet1.InsertRow(currentRow, 1); // 插入一行
                    sheet1.Row(currentRow).Height = 28.5;
                    sheet1.Cells[currentRow, 3, currentRow, 4].Merge = true;

                    for (int col = 1; col <= 10; col++)
                    {
                        sheet1.Cells[currentRow, col].StyleID = sheet1.Cells[currentRow - 1, col].StyleID;
                    }
                }

                sheet1.Cells[currentRow, 2].Value = currentRow - startRow + 1;
                sheet1.Cells[currentRow, 3].Value = Item.ItemName;
                sheet1.Cells[currentRow, 5].Value = Item.Unit;
                sheet1.Cells[currentRow, 6].Value = Item.UnitPrice;
                sheet1.Cells[currentRow, 7].Value = Item.Quantity;
                sheet1.Cells[currentRow, 8].Value = Item.TotalAmount;
                sheet1.Cells[currentRow, 9].Value = Item.Note;

                currentRow++;
            }

            if (currentRow < 15) currentRow = 15;

            sheet1.Cells[currentRow, 8].Value = totalAmount;
            sheet1.Cells[currentRow + 1, 8].Value = tax;
            sheet1.Cells[currentRow + 2, 8].Value = totalWithTax;

            package.SaveAs(new FileInfo(newFileName));
        }
    }
}
