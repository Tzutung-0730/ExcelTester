using System;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelCertificateBill
    {
        public static IEnumerable<ExcelColumnCertificateBill> GetData()
        {
            var certificateBills = new List<ExcelColumnCertificateBill>
            {
                new ExcelColumnCertificateBill
                {
                    CertificateBillId = Guid.NewGuid(), // 憑證帳單Id
                    BillYear = 2022, // 帳單年份
                    BillMonth = 11, // 帳單月份
                    PaymentDeadline = new DateTime(2022, 12, 30), // 付款截止日
                    SettleTime = new DateTime(2022, 12, 15), // 結帳日
                    ModifyTime = DateTime.Now, // 修改時間
                    Modifier = Guid.NewGuid(), // 修改者 ID
                    Items = new List<ExcelColumnCertificateBillItem>
                    {
                        new ExcelColumnCertificateBillItem
                        {
                            CertificateBillItemId = Guid.NewGuid(), // 憑證帳單項目Id
                            CertificateBillId = Guid.NewGuid(), // 憑證帳單Id
                            ProjectId = "C001", // 專案Id
                            SupplierPlaceId = Guid.NewGuid(), // 供應商地點Id
                            SupplierPlaceName = "苗栗竹南風場憑證", // 供應商地點名稱
                            CertificateCount = 1246, // 憑證數量
                            CertificateFee = 3738, // 憑證費用
                            ServiceFee = 0, // 服務費用
                            ModifyTime = DateTime.Now, // 修改時間
                            Modifier = Guid.NewGuid(), // 修改者
                            IsExtra = false // 是否為附加費用
                        },
                        new ExcelColumnCertificateBillItem
                        {
                            CertificateBillItemId = Guid.NewGuid(),
                            CertificateBillId = Guid.NewGuid(),
                            ProjectId = "C002",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "苗栗大鵬風場憑證",
                            CertificateCount = 3438,
                            CertificateFee = 10314,
                            ServiceFee = 0,
                            ModifyTime = DateTime.Now,
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        },
                        new ExcelColumnCertificateBillItem
                        {
                            CertificateBillItemId = Guid.NewGuid(),
                            CertificateBillId = Guid.NewGuid(),
                            ProjectId = "C004",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "鑫光憑證",
                            CertificateCount = 381,
                            CertificateFee = 1143,
                            ServiceFee = 155,
                            ModifyTime = DateTime.Now,
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        },
                        new ExcelColumnCertificateBillItem
                        {
                            CertificateBillItemId = Guid.NewGuid(),
                            CertificateBillId = Guid.NewGuid(),
                            ProjectId = "C006",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "嘉南台南頭水力電廠憑證",
                            CertificateCount = 2489,
                            CertificateFee = 7467,
                            ServiceFee = 0,
                            ModifyTime = DateTime.Now,
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        },
                        new ExcelColumnCertificateBillItem
                        {
                            CertificateBillItemId = Guid.NewGuid(),
                            CertificateBillId = Guid.NewGuid(),
                            ProjectId = "C007",
                            SupplierPlaceId = Guid.NewGuid(),
                            SupplierPlaceName = "華祥力麗憑證",
                            CertificateCount = 220,
                            CertificateFee = 660,
                            ServiceFee = 105,
                            ModifyTime = DateTime.Now,
                            Modifier = Guid.NewGuid(),
                            IsExtra = false
                        }
                    }
                }
            };

            return certificateBills;
        }
    }
}
