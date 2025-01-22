using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelTester.Classes
{
    internal class SupplierContractBillWord
    {
        public static void WriteWord(string srcTemplateFile, string newFileName, IEnumerable<ExcelColumnSupplierContractBill> excelData)
        {
            var tmpContent = File.ReadAllBytes(srcTemplateFile);
            using var ms = new MemoryStream();
            ms.Write(tmpContent, 0, tmpContent.Length);
            var bill = excelData.First();

            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                if (doc.MainDocumentPart == null)
                {
                    doc.AddMainDocumentPart();
                    doc.MainDocumentPart.Document = new Document(new Body());
                }

                var body = doc.MainDocumentPart.Document.Body;

                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center }), // 置中
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }, // 字型
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "32" } // 字型大小：16 號 * 2
                        ),
                        new Text("台汽綠能股份有限公司"),
                        new Break(),
                        new Text("驗收紀錄表")
                    ),
                    new Paragraph()
                ));

                // 設定邊界：上下左右各 2 公分 (1 公分 = 567 點，2 公分 = 1134 點)
                body.Append(new SectionProperties(
                    new PageMargin
                    {
                        Top = 1134, // 上邊界
                        Bottom = 1134, // 下邊界
                        Left = 1134, // 左邊界
                        Right = 1134 // 右邊界
                    }
                ));

                var table1 = new Table();
                table1.AppendChild(new TableProperties(
                    new TableIndentation { Width = 360, Type = TableWidthUnitValues.Dxa }, // 向右縮進 720 pt (4 格)
                    new TableBorders(
                        new TopBorder { Val = BorderValues.None },
                        new BottomBorder { Val = BorderValues.None },
                        new LeftBorder { Val = BorderValues.None },
                        new RightBorder { Val = BorderValues.None },
                        new InsideHorizontalBorder { Val = BorderValues.None },
                        new InsideVerticalBorder { Val = BorderValues.None }
                        )
                ));

                var row1_1 = new TableRow();
                row1_1.Append(new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "5200" } // 設置單元格寬度
                    ),
                    new Paragraph(
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text($"驗收日期：{bill.SettleTime.Year - 1911}/{bill.SettleTime:MM/dd}")
                        )
                    )
                ));
                row1_1.Append(new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4000" } // 設置單元格寬度
                    ),
                    new Paragraph(
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text($"廠商名稱：{bill.SupplierName}")
                        )
                    )
                ));
                table1.Append(row1_1);

                var row1_2 = new TableRow();
                row1_2.Append(new TableCell(new Paragraph(new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                    ),
                    new Text("合約編號及名稱：再生能源電能購售契約書")
                ))));
                row1_2.Append(new TableCell(new Paragraph(new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                    ),
                    new Text("承包金額：")
                ))));
                table1.Append(row1_2);

                var row1_3 = new TableRow();
                row1_3.Append(new TableCell(new Paragraph(new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                    ),
                    new Text("驗收內容及結果：")
                ))));
                row1_3.Append(new TableCell(new Paragraph(new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                    ),
                    new Text("")
                ))));
                table1.Append(row1_3);

                body.Append(table1);

                body.Append(new Paragraph(
                    new ParagraphProperties(),
                    new Run(
                        new RunProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "2" } // 字體大小設置為 1 號字
                        ),
                        new Text("\u2009")
                    )
                ));

                var table2 = new Table();
                table2.AppendChild(new TableProperties(
                    new TableIndentation { Width = 360, Type = TableWidthUnitValues.Dxa }, // 向右縮進 360 pt (2 格)
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 8 },
                        new BottomBorder { Val = BorderValues.Single, Size = 8 },
                        new LeftBorder { Val = BorderValues.Single, Size = 8 },
                        new RightBorder {Val = BorderValues.Single, Size = 8 },
                        new InsideHorizontalBorder {Val = BorderValues.Single, Size = 8 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 8 }
                        )
                ));

                var row2_1 = new TableRow(
                    new TableRowProperties(
                        new TableRowHeight { Val = 731 } // 設定行高為 731 Twips
                    )
                );
                row2_1.Append(new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "800" },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center } 
                    ), 
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text("項次")
                        )
                    )
                ));
                row2_1.Append(new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1500" },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center } 
                    ),
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text("項目規格")
                        )
                    )
                ));
                row2_1.Append(new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "900" },
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center } 
                    ),
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text("數量")
                        )
                    )
                ));
                row2_1.Append(new TableCell(
                    new TableCellProperties(
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                    ),
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text("驗  收  內  容")
                        )
                    )
                ));
                table2.Append(row2_1);

                var startDate = new DateTime(bill.BillYear, bill.BillMonth, 1);
                var endDate = startDate.AddMonths(1).AddDays(-1);
                var supplierElectricityNumber = string.Join("、", bill.Items
                    .Where(i => !string.IsNullOrEmpty(i.SupplierElectricityNumber)) // 過濾掉空值
                    .Select(i => i.SupplierElectricityNumber)
                    .Distinct()); // 過濾重複值

                var row2_2 = new TableRow(
                    new TableRowProperties(
                        new TableRowHeight { Val = 731 }
                    )
                );
                row2_2.Append(new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text("1.")
                        )
                    )
                ));
                row2_2.Append(new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Left }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text("再生能源發電費"),
                            new Break(),
                            new Text($"({startDate.Year - 1911}/{startDate:MM/dd}~{endDate.Year - 1911}/{endDate:MM/dd})")
                        )
                    )
                ));
                row2_2.Append(new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Center }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text((bill.Items.Sum(x => x.Quantity ?? 0)).ToString("N0"))
                        )
                    )
                ));
                row2_2.Append(new TableCell(
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification { Val = JustificationValues.Left }
                        ),
                        new Run(
                            new RunProperties(
                                new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" }
                            ),
                            new Text($"{bill.BillYear - 1911}/{bill.BillMonth:D2} 本公司購電度數總計為「調整後-當月總用電量(瓩)」度，再生能源電能費用未稅金額為 NT${Math.Round(bill.Items.Sum(i => i.TotalAmount * 0.95)):N0}元，含稅金額為 NT${bill.Items.Sum(i => i.TotalAmount):N0}元。"),
                            new Break(),
                            new Text("經承辦人員核對確認無誤。"),
                            new Break(),
                            new Text($"案場電號：{supplierElectricityNumber}")
                        )
                    )
                ));
                table2.Append(row2_2);
                
                body.Append(table2);
                
                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new Justification { Val = JustificationValues.Left },
                        new SpacingBetweenLines
                        {
                            Line = "480",
                            Before = "240" // 段落前距離，單位為 Twips (1/20 點)
                        }
                    ),
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "24" } 
                        ),
                        new Text("   ※本欄不適用或不足時，請另以附件列表說明") { Space = SpaceProcessingModeValues.Preserve },
                        new Break(),
                        new Text("   4. 驗收意見：合格■，不合格□，要求改善□，其他□") { Space = SpaceProcessingModeValues.Preserve },
                        new Break(),
                        new Text($"   5. 驗收合格日：{bill.SettleTime.Year - 1911}/{bill.SettleTime:MM/dd}") { Space = SpaceProcessingModeValues.Preserve },
                        new Break(),
                        new Text("   6. 附件： 台電繳款通知單、再生能源電能購售契約書") { Space = SpaceProcessingModeValues.Preserve },
                        new Break(),
                        new Text("   7. 其他：") { Space = SpaceProcessingModeValues.Preserve }
                    )
                ));

                var table3 = new Table();
                table3.AppendChild(new TableProperties(
                    new TableWidth { Width = "10000", Type = TableWidthUnitValues.Dxa }, // 表格總寬度
                    new TableIndentation { Width = 360, Type = TableWidthUnitValues.Dxa },
                    new TableBorders( 
                        new TopBorder { Val = BorderValues.None },
                        new BottomBorder { Val = BorderValues.None },
                        new LeftBorder {Val = BorderValues.None },
                        new RightBorder { Val = BorderValues.None },
                        new InsideHorizontalBorder {Val = BorderValues.None },
                        new InsideVerticalBorder {Val = BorderValues.None }
                    )
                ));
                var row3_1 = new TableRow();
                string[] verticalTexts = { "核定", "主管", "監驗", "會驗", "主驗", "經辦" };
                foreach (var verticalText in verticalTexts)
                {
                    var verticalRun = new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "24" }
                        )
                    );
                    foreach (var ch in verticalText)
                    {
                        verticalRun.Append(new Text(ch.ToString()));
                        verticalRun.Append(new Break());
                    }

                    row3_1.Append(new TableCell(
                        new TableCellProperties(
                            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                        ),
                        new Paragraph(
                            new ParagraphProperties(
                                new Justification { Val = JustificationValues.Left }
                            ),
                            verticalRun
                        )
                    ));
                }
                table3.Append(row3_1);
                body.Append(table3);

                body.Append(new Paragraph(
                    new ParagraphProperties(
                        new SpacingBetweenLines
                        {
                            Before = "720" 
                        }
                    ),
                    new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "標楷體", EastAsia = "標楷體" },
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "24" }
                        ),
                        new Break(),
                        new Text("   表單編號：TGE-PS-S02") { Space = SpaceProcessingModeValues.Preserve }
                    )
                ));

                doc.MainDocumentPart.Document.Save();
            }

            File.WriteAllBytes(newFileName, ms.ToArray());
        }
    }
}
