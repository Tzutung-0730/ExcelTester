using System;
using System.Collections.Generic;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelCustomerContractBill
    {
        public static IEnumerable<ExcelColumnCustomerContractBill> GetData()
        {
            var customerContractBills = new List<ExcelColumnCustomerContractBill>
            {
                new ExcelColumnCustomerContractBill
                {
                    CustomerContractBillId = Guid.NewGuid(), // 客戶契約帳單唯一識別碼
                    CustomerId = Guid.NewGuid(), // 客戶唯一識別碼
                    CustomerName = "台灣優衣庫有限公司", // 客戶名稱
                    CustomerContractId = Guid.NewGuid(), // 客戶契約唯一識別碼
                    CustomerTaxIDNumber = "25117554", // 客戶統一編號
                    CustomerAddress = "新北市板橋區縣民大道2段68號15樓", // 客戶地址
                    CustomerContactNumber = "02-8953-2486", // 客戶聯絡電話
                    BillCode = "20230413", // 帳單編號
                    BillYear = 2022, // 帳單年份
                    BillMonth = 7, // 帳單月份
                    BillingMethod = BillingMethodType.GroupByBelong, // 帳單開立方式
                    ReceiveItem = ReceiveItemType.ElecRedistributeCertAuditCertServ, // 收款項目
                    IsSplitByReceiveItem = false, // 是否依收款項目拆分帳單
                    IndustCategory = "農業", // 產業類別
                    EelecBelong = "台灣電綜", // 電號所屬
                    PaymentDeadline = new DateTime(2023, 8, 25), // 付款截止日期
                    TotalPowerUse = 229290, // 當月總用電量
                    SettleTime = new DateTime(2023, 4, 13), // 結帳日期
                    ModifyTime = DateTime.Now, // 修改時間
                    Modifier = Guid.NewGuid(), // 修改者
                    Items = new List<ExcelColumnCustomerContractBillItem>
                    {
                        new ExcelColumnCustomerContractBillItem
                        {
                            CustomerContractBillItemId = Guid.NewGuid(), // 帳單項目唯一識別碼
                            CustomerContractBillId = Guid.NewGuid(), // 關聯帳單唯一識別碼
                            SortOrder = 1, // 排序
                            ItemName = "再生能源發電費", // 項目名稱
                            ProjectId = "C007", // 案場專案代號
                            Unit = "度", // 單位
                            UnitPrice = 5.7500m, // 單價
                            Quantity = 229290, // 數量
                            TotalAmount = 1318418, // 總金額
                            Note = "", // 備註
                            Modifier = Guid.NewGuid(), // 修改者
                            IsExtra = false // 是否為附加費用
                        },
                        new ExcelColumnCustomerContractBillItem
                        {
                            CustomerContractBillItemId = Guid.NewGuid(),
                            CustomerContractBillId = Guid.NewGuid(),
                            SortOrder = 2,
                            ItemName = "電代購憑證服務費用",
                            ProjectId = "C007",
                            Unit = "式",
                            UnitPrice = 12009,
                            Quantity = 1,
                            TotalAmount = 12009,
                            Note = "",
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        },
                        new ExcelColumnCustomerContractBillItem
                        {
                            CustomerContractBillItemId = Guid.NewGuid(),
                            CustomerContractBillId = Guid.NewGuid(),
                            SortOrder = 3,
                            ItemName = "再生能源憑證登錄費",
                            ProjectId = "C007",
                            Unit = "張",
                            UnitPrice = 690,
                            Quantity = 1,
                            TotalAmount = 690,
                            Note = "",
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        },
                        new ExcelColumnCustomerContractBillItem
                        {
                            CustomerContractBillItemId = Guid.NewGuid(),
                            CustomerContractBillId = Guid.NewGuid(),
                            SortOrder = 4,
                            ItemName = "再生能源憑證服務費",
                            ProjectId = "C007",
                            Unit = "張",
                            UnitPrice = 111,
                            Quantity = 1,
                            TotalAmount = 111,
                            Note = "",
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        }
                    }
                }
            };

            return customerContractBills;
        }
    }
}
