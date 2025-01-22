using System;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelTaipowerBill
    {
        public static IEnumerable<ExcelColumnTaipowerBill> GetData()
        {
            var taipowerBills = new List<ExcelColumnTaipowerBill>
            {
                new ExcelColumnTaipowerBill
                {
                    TaipowerBillId = Guid.NewGuid(), // 台電帳單 ID
                    BillYear = 2022, // 帳單年份
                    BillMonth = 12, // 帳單月份
                    PaymentDeadline = new DateTime(2023, 2, 1), // 付款截止日
                    SettleTime = new DateTime(2023, 1, 18), // 結帳日
                    Modifier = Guid.NewGuid(), // 修改者 ID
                    Items = new List<ExcelColumnTaipowerBillItem>
                    {
                        // C001 專案代號
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C001",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "苗栗風力-竹南",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台積電 P6",
                            TotalAmount = 161738,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C001",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "苗栗風力-大鵬",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台積電 P6",
                            TotalAmount = 574537,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C001",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "苗栗風力-大鵬",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台積電 P4",
                            TotalAmount = 143634,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        // C004 專案代號
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C004",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "鑫光",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台灣固鋼",
                            TotalAmount = 12297,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C004",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "鑫光",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台綜院",
                            TotalAmount = 117,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C004",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "鑫光",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "富邦銀行",
                            TotalAmount = 1842,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C004",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "鑫光",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "櫃買中心",
                            TotalAmount = 1006,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        // C005 專案代號
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C005",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "星寶",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台達電",
                            TotalAmount = 196245,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C005",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "星寶",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台固網及台灣大",
                            TotalAmount = 99105,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        },
                        // C007 專案代號
                        new ExcelColumnTaipowerBillItem
                        {
                            TaipowerBillItemId = Guid.NewGuid(),
                            TaipowerBillId = Guid.NewGuid(),
                            ProjectId = "C007",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "華洋力麗",
                            CustomerId = Guid.NewGuid(),
                            CustomerName = "台灣優衣大",
                            TotalAmount = 12609,
                            IsExtra = false,
                            Modifier = Guid.NewGuid()
                        }
                    }
                }
            };

            return taipowerBills;
        }
    }
}
