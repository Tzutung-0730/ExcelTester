using Bogus;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelTaipowerBill
    {
        // 生成假資料 for TaipowerBill（包含 Items 資料）
        public static IEnumerable<ExcelColumnTaipowerBill> GenerateFakeData()
        {
            // 生成單一 TaipowerBill 的假資料
            var billFaker = new Faker<ExcelColumnTaipowerBill>()
                .RuleFor(o => o.TaipowerBillId, f => Guid.NewGuid()) // 台電帳單 ID
                .RuleFor(o => o.BillYear, f => f.Date.Past(10).Year) // 帳單年份
                .RuleFor(o => o.BillMonth, f => f.Date.Past(10).Month) // 帳單月份
                .RuleFor(o => o.PaymentDeadline, f => f.Date.Future()) // 付款截止日
                .RuleFor(o => o.SettleTime, f => f.Date.Future()) // 結帳日
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()) // 修改者 ID
                .RuleFor(o => o.Items, f => GenerateFakeDataForItems().Take(f.Random.Int(1, 5)).ToList()); // 帳單項目清單

            return billFaker.Generate(1); // 生成 1 筆 TaipowerBill 假資料
        }

        // 生成假資料 for TaipowerBillItem
        public static IEnumerable<ExcelColumnTaipowerBillItem> GenerateFakeDataForItems()
        {
            var faker = new Faker<ExcelColumnTaipowerBillItem>()
                .RuleFor(o => o.ProjectId, f => f.Random.AlphaNumeric(10)) // 專案Id
                .RuleFor(o => o.SupplierPlaceName, f => f.Company.CompanyName()) // 供應商地點名稱
                .RuleFor(o => o.CustomerName, f => f.Name.FullName()) // 客戶名稱
                .RuleFor(o => o.TotalAmount, f => f.Random.Int(1000, 50000)) // 總金額
                .RuleFor(o => o.IsExtra, f => f.Random.Bool()) // 是否為附加費用
                .RuleFor(o => o.TaipowerBillItemId, f => Guid.NewGuid()) // 台電帳單項目 ID
                .RuleFor(o => o.TaipowerBillId, f => Guid.NewGuid()) // 台電帳單 ID
                .RuleFor(o => o.SupplierPlaceId, f => Guid.NewGuid()) // 供應商地點 ID
                .RuleFor(o => o.CustomerId, f => Guid.NewGuid()) // 客戶 ID
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()); // 修改者 ID

            return faker.Generate(50); // 生成 50 筆 TaipowerBillItem 假資料
        }
    }
}
