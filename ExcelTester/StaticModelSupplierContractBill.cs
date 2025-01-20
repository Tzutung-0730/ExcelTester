using Bogus;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModelSupplierContractBill
    {
        public static IEnumerable<ExcelColumnSupplierContractBill> GenerateFakeData()
        {
            // 使用 Faker 生成假資料
            var billFaker = new Faker<ExcelColumnSupplierContractBill>()
                .RuleFor(o => o.SupplierContractBillId, f => Guid.NewGuid()) // 帳單唯一識別碼
                .RuleFor(o => o.SupplierId, f => Guid.NewGuid()) // 電業公司唯一識別碼
                .RuleFor(o => o.SupplierName, f => f.Random.AlphaNumeric(8)) // 電業公司名稱
                .RuleFor(o => o.SupplierContractId, f => Guid.NewGuid()) // 電業公司契約唯一識別碼
                .RuleFor(o => o.SupplierTaxIDNumber, f => f.Random.String2(8, "0123456789")) // 電業公司統一編號
                .RuleFor(o => o.SupplierAddress, f => f.Address.FullAddress()) // 電業公司地址
                .RuleFor(o => o.SupplierContactNumber, f => f.Phone.PhoneNumber()) // 電業公司聯絡電話
                .RuleFor(o => o.BillCode, f => f.Commerce.Ean13()) // 帳單編號
                .RuleFor(o => o.BillYear, f => f.Date.Past(1).Year) // 帳單年度
                .RuleFor(o => o.BillMonth, f => f.Date.Past(1).Month) // 帳單月份
                .RuleFor(o => o.PaymentDeadline, f => f.Date.Future()) // 付款截止日
                .RuleFor(o => o.PaymentMethod, f => f.Finance.TransactionType()) // 付款方式
                .RuleFor(o => o.ReceiveAccount, f => f.Finance.Account()) // 收款帳戶
                .RuleFor(o => o.ReceiveName, f => f.Name.FullName()) // 收款帳戶名稱
                .RuleFor(o => o.ReceiveBankName, f => f.Company.CompanyName()) // 收款帳戶銀行名稱
                .RuleFor(o => o.SettleTime, f => f.Date.Future()) // 結算時間
                .RuleFor(o => o.ModifyTime, f => f.Date.Recent()) // 修改時間
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()); // 修改者

            // 生成 1 筆 SupplierContractBill 假資料
            var supplierContractBill = billFaker.Generate(1).First();

            // 創建 Items 並逐筆添加
            var itemFaker = new Faker<ExcelColumnSupplierContractBillItem>()
                .RuleFor(o => o.SupplierContractBillItemId, f => Guid.NewGuid()) // 帳單項目唯一識別碼
                .RuleFor(o => o.SupplierContractBillId, f => Guid.NewGuid()) // 關聯帳單唯一識別碼
                .RuleFor(o => o.SortOrder, f => f.Random.Int(1, 10)) // 排序
                .RuleFor(o => o.ItemName, f => f.Commerce.ProductName()) // 項目名稱
                .RuleFor(o => o.SupplierPlaceId, f => Guid.NewGuid()) // 電業公司案場唯一識別碼
                .RuleFor(o => o.SupplierPlaceName, f => f.Address.City()) // 案場名稱
                .RuleFor(o => o.SupplierElectricityNumber, f => f.Random.AlphaNumeric(8)) // 電號
                .RuleFor(o => o.ProjectId, f => f.Random.AlphaNumeric(5)) // 案場專案代號
                .RuleFor(o => o.Unit, f => f.Commerce.ProductMaterial()) // 單位
                .RuleFor(o => o.UnitPrice, f => f.Random.Decimal(10, 500)) // 單價
                .RuleFor(o => o.Quantity, f => f.Random.Int(1, 100)) // 數量
                .RuleFor(o => o.TotalAmount, (f, o) => (int)(o.UnitPrice.GetValueOrDefault() * o.Quantity.GetValueOrDefault())) // 總金額
                .RuleFor(o => o.Note, f => f.Lorem.Sentence()) // 備註
                .RuleFor(o => o.ModifyTime, f => f.Date.Recent()) // 修改時間
                .RuleFor(o => o.Modifier, f => Guid.NewGuid()) // 修改者
                .RuleFor(o => o.IsExtra, f => f.Random.Bool()); // 是否為附加費用

            // 創建 Items 並逐筆添加到 SupplierContractBill 的 Items 屬性
            var items = new List<ExcelColumnSupplierContractBillItem>();
            var itemCount = new Faker().Random.Int(2, 5); // 隨機生成 2 到 5 個項目
            for (int i = 0; i < itemCount; i++)
            {
                items.Add(itemFaker.Generate());
            }

            supplierContractBill.Items = items; // 將 Items 資料加入 SupplierContractBill

            return new List<ExcelColumnSupplierContractBill> { supplierContractBill }; // 返回包含 1 筆假資料的列表
        }
    }
}
