using Bogus;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelCertificateBill
    {
        public static IEnumerable<ExcelColumnCertificateBill> GenerateFakeData()
        {
            // 使用 Faker 生成假資料
            var billFaker = new Faker<ExcelColumnCertificateBill>()
                .RuleFor(o => o.CertificateBillId, f => Guid.NewGuid()) // 憑證帳單Id
                .RuleFor(o => o.BillYear, f => f.Date.Past(10).Year) // 帳單年份
                .RuleFor(o => o.BillMonth, f => f.Date.Past(10).Month) // 帳單月份
                .RuleFor(o => o.PaymentDeadline, f => f.Date.Future()) // 付款截止日
                .RuleFor(o => o.SettleTime, f => f.Date.Future()) // 結帳日
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()); // 修改者 ID

            // 生成 1 筆 CertificateBill 假資料
            var certificateBill = billFaker.Generate(1).First();

            // 創建 Items 並逐筆添加
            var itemFaker = new Faker<ExcelColumnCertificateBillItem>()
                .RuleFor(o => o.CertificateBillItemId, f => Guid.NewGuid()) // 憑證帳單項目Id
                .RuleFor(o => o.CertificateBillId, f => Guid.NewGuid()) // 憑證帳單Id
                .RuleFor(o => o.ProjectId, f => f.Random.AlphaNumeric(10)) // 專案Id
                .RuleFor(o => o.SupplierPlaceId, f => f.Random.Guid()) // 供應商地點Id
                .RuleFor(o => o.SupplierPlaceName, f => f.Company.CompanyName()) // 供應商地點名稱
                .RuleFor(o => o.CertificateCount, f => f.Random.Int(1, 100)) // 憑證數量
                .RuleFor(o => o.CertificateFee, f => f.Random.Int(1000, 5000)) // 憑證費用
                .RuleFor(o => o.ServiceFee, f => f.Random.Int(500, 2000)) // 服務費用
                .RuleFor(o => o.ModifyTime, f => f.Date.Recent()) // 修改時間
                .RuleFor(o => o.Modifier, f => f.Random.Guid()) // 修改者
                .RuleFor(o => o.IsExtra, f => f.Random.Bool()); // 是否為附加費用

            // 創建 Items 並逐筆添加到 CertificateBill 的 Items 屬性
            var items = new List<ExcelColumnCertificateBillItem>();
            var itemCount = new Faker().Random.Int(2, 5); // 隨機生成 2 到 5 個項目
            for (int i = 0; i < itemCount; i++)
            {
                items.Add(itemFaker.Generate());
            }

            certificateBill.Items = items; // 將 Items 資料加入 CertificateBill

            return new List<ExcelColumnCertificateBill> { certificateBill }; // 返回包含 1 筆假資料的列表
        }
    }
}
