using System;
using System.Collections.Generic;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelSupplierContractBill
    {
        public static IEnumerable<ExcelColumnSupplierContractBill> GetData()
        {
            var supplierContractBill = new List<ExcelColumnSupplierContractBill>
            {
                new ExcelColumnSupplierContractBill
                {
                    SupplierContractBillId = Guid.NewGuid(), // 唯一識別碼
                    SupplierId = Guid.NewGuid(), // 電業公司唯一識別碼
                    SupplierName = "苗栗風力股份有限公司", // 電業公司名稱
                    SupplierContractId = Guid.NewGuid(), // 電業公司契約唯一識別碼
                    SupplierTaxIDNumber = "80520735", // 電業公司統一編號
                    SupplierAddress = "台北市內湖區瑞光路392號6F", // 電業公司地址
                    SupplierContactNumber = "02-8798-2899 #610", // 電業公司聯絡電話
                    BillCode = "MW-0001-2212", // 帳單編號
                    BillYear = 2022, // 帳單年度
                    BillMonth = 12, // 帳單月份
                    PaymentDeadline = new DateTime(2023, 3, 6), // 付款截止日
                    PaymentMethod = "電匯", // 付款方式
                    ReceiveAccount = "20609001655", // 收款帳戶
                    ReceiveName = "苗栗風力股份有限公司", // 收款帳戶名稱
                    ReceiveBankName = "兆豐國際商業銀行 板橋分行", // 收款帳戶銀行名稱
                    SettleTime = new DateTime(2025, 1, 21), // 結算時間
                    ModifyTime = DateTime.Now, // 修改時間
                    Modifier = Guid.NewGuid(), // 修改者
                    Items = new List<ExcelColumnSupplierContractBillItem> // 帳單項目列表
                    {
                        new ExcelColumnSupplierContractBillItem
                        {
                            SupplierContractBillItemId = Guid.NewGuid(), // 帳單項目唯一識別碼
                            SupplierContractBillId = Guid.NewGuid(), // 關聯帳單唯一識別碼
                            SortOrder = 1, // 排序
                            ItemName = "再生能源憑證電費", // 項目名稱
                            SupplierPlaceId = Guid.NewGuid(), // 電業公司案場唯一識別碼
                            SupplierPlaceName = "苗栗風力股份有限公司", // 案場名稱
                            SupplierElectricityNumber = "20885560111", // 電號
                            ProjectId = "C001", // 案場專案代號
                            Unit = "度", // 單位
                            UnitPrice = 3.175m, // 單價
                            Quantity = 21740992, // 數量
                            TotalAmount = 69027650, // 金額（未稅總金額）
                            Note = "", // 備註
                            ModifyTime = DateTime.Now, // 修改時間
                            Modifier = Guid.NewGuid(), // 修改者
                            IsExtra = false // 是否為附加費用
                        }
                    }
                }
            };

            return supplierContractBill;
        }
    }
}
