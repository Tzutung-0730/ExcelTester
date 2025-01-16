using ExcelTester.Attributes;
using OfficeOpenXml;

namespace ExcelTester.Classes
{
        public class Report
        {
            /// <summary>
            /// 產生報表檔案
            /// </summary>
            /// <returns></returns>
            public async Task GeneralReport(string srcTemplateFile, string newFileName)
            {
                // 產生資料
                // var report = await QueryReport();

                // 寫入 Excel
                WriteExcel(srcTemplateFile, newFileName, StaticModel.GenerateFakeData());
            }

            /// <summary>
            /// 寫入 Excel
            /// </summary>
            /// <param name="excelData"></param>
            public void WriteExcel(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnReport> excelData)
        {
            // Set the ExcelPackage LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 讀取範本檔
            var package = new ExcelPackage(new FileInfo(srcTemplateFile));

                ExcelWorksheet sheet = package.Workbook.Worksheets[0];

                // 停止繪畫的功能
                package.DoAdjustDrawings = false;

                // 第一列是表頭要跳過
                int currentRow = 2;

                // 將資料寫入 Excel  
                foreach (var item in excelData)
                {
                    ItemToRow(sheet, currentRow, item);
                    currentRow++;
                }

                // 計算 PivotTable 樞紐分析表
                foreach (var pivotSheet in package.Workbook.Worksheets)
                {
                    foreach (var pivotTable in pivotSheet.PivotTables)
                    {
                        // 對指定的樞紐分析表進行排序
                        // 但目前有亂排的問題，先註解掉
                        //if (pivotSheet.Name == "長條圖-07" || pivotSheet.Name == "長條圖-08")
                        //{
                        //    pivotTable.ColumnFields[0].Sort = eSortType.None;
                        //    pivotTable.ColumnFields[0].AddNumericGrouping(0, 150, 10);
                        //}
                        pivotTable.Calculate(true);
                    }
                }

                // 恢復繪畫的功能
                package.DoAdjustDrawings = true;

                // 用 GUID 產生一個隨機的檔名
                package.SaveAs(newFileName);
            }

            /// <summary>
            /// 將 Item 變成 Excel 的 Row
            /// </summary>
            /// <param name="sheet"></param>
            /// <param name="row"></param>
            /// <param name="item"></param>
            private static void ItemToRow(ExcelWorksheet sheet, int row, ExcelColumnReport item)
            {
                foreach (var prop in item.GetType().GetProperties())
                {
                    ExcelColumnMapping mapping = (ExcelColumnMapping)Attribute.GetCustomAttribute(prop, typeof(ExcelColumnMapping));

                    if (mapping != null)
                    {
                        int col = mapping.ColumnIndex;
                        sheet.Cells[row, col].Value = prop.GetValue(item);
                    }
                }
            }
        }
    }
