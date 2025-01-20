using Bogus;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelCustomerContractBill
    {
        // 生成假資料 for TaipowerBill（包含 Items 資料）
        public static IEnumerable<ExcelColumnCustomerContractBill> GenerateFakeData()
        {
            // 生成單一 TaipowerBill 的假資料
            var billFaker = new Faker<ExcelColumnCustomerContractBill>()
                .RuleFor(o => o.CustomerContractBillId, f => Guid.NewGuid()) // 客戶契約帳單唯一識別碼
                .RuleFor(o => o.CustomerId, f => Guid.NewGuid()) // 客戶唯一識別碼
                .RuleFor(o => o.CustomerName, f => f.Company.CompanyName()) // 客戶名稱
                .RuleFor(o => o.CustomerContractId, f => Guid.NewGuid()) // 客戶契約唯一識別碼
                .RuleFor(o => o.CustomerTaxIDNumber, f => f.Random.String2(8, "0123456789")) // 客戶統一編號
                .RuleFor(o => o.CustomerAddress, f => f.Address.FullAddress()) // 客戶地址
                .RuleFor(o => o.CustomerContactNumber, f => f.Phone.PhoneNumber()) // 客戶聯絡電話
                .RuleFor(o => o.BillCode, f => f.Commerce.Ean13()) // 帳單編碼
                .RuleFor(o => o.BillYear, f => f.Date.Past(5).Year) // 帳單年份
                .RuleFor(o => o.BillMonth, f => f.Date.Past(1).Month) // 帳單月份
                .RuleFor(o => o.BillingMethod, f => f.PickRandom<BillingMethodType>()) // 帳單開立方式
                .RuleFor(o => o.ReceiveItem, f => f.PickRandom<ReceiveItemType>()) // 收款項目
                .RuleFor(o => o.IsSplitByReceiveItem, f => f.Random.Bool()) // 是否依收款項目拆分帳單
                .RuleFor(o => o.IndustCategory, f => f.Commerce.Department()) // 產業類別
                .RuleFor(o => o.EelecBelong, f => f.Company.CompanyName()) // 電號所屬
                .RuleFor(o => o.PaymentDeadline, f => f.Date.Future(1)) // 付款截止日期
                .RuleFor(o => o.TotalPowerUse, f => f.Random.Int(1000, 100000)) // 當月總用電量
                .RuleFor(o => o.SettleTime, f => f.Date.Future(1)) // 結帳日期
                .RuleFor(o => o.ModifyTime, f => f.Date.Recent()) // 修改時間
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()); // 修改者

            // 生成 1 筆 TaipowerBill 假資料
            var customerContractBill = billFaker.Generate(1).First();

            // 創建 Items 並逐筆添加
            var itemFaker = new Faker<ExcelColumnCustomerContractBillItem>()
                .RuleFor(o => o.CustomerContractBillItemId, f => Guid.NewGuid()) // 帳單項目Id
                .RuleFor(o => o.CustomerContractBillId, f => Guid.NewGuid()) // 帳單Id
                .RuleFor(o => o.SortOrder, f => f.Random.Int(1, 10)) // 排序
                .RuleFor(o => o.ItemName, f => f.Commerce.ProductName()) // 項目名稱
                .RuleFor(o => o.ProjectId, f => f.Random.AlphaNumeric(10)) // 案場專案代號
                .RuleFor(o => o.Unit, f => f.Random.ArrayElement(new[] { "kWh", "MW", "GWh" })) // 單位
                .RuleFor(o => o.UnitPrice, f => f.Random.Decimal(1, 1000)) // 單價
                .RuleFor(o => o.Quantity, f => f.Random.Int(1, 100)) // 數量
                .RuleFor(o => o.TotalAmount, (f, o) => (int)((o.UnitPrice ?? 0) * (o.Quantity ?? 0))) // 總金額
                .RuleFor(o => o.Note, f => f.Lorem.Sentence()) // 備註
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()) // 修改者
                .RuleFor(o => o.IsExtra, f => f.Random.Bool()); // 是否為附加費用

            // 創建 Items 並逐筆添加到 TaipowerBill 的 Items 屬性
            var items = new List<ExcelColumnCustomerContractBillItem>();
            var itemCount = new Faker().Random.Int(2, 5); // 隨機生成 2 到 5 個項目
            for (int i = 0; i < itemCount; i++)
            {
                items.Add(itemFaker.Generate());
            }

            customerContractBill.Items = items; // 將 Items 資料加入 TaipowerBill

            return new List<ExcelColumnCustomerContractBill> { customerContractBill }; // 返回包含 1 筆假資料的列表
        }
    }
}
