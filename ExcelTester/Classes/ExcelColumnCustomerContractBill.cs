namespace ExcelTester.Classes
{
    public class ExcelColumnCustomerContractBill
    {
        public Guid CustomerContractBillId { get; set; }
        public Guid CustomerId { get; set; }
        public required string CustomerName { get; set; }
        public Guid CustomerContractId { get; set; }
        public required string CustomerTaxIDNumber { get; set; }
        public required string CustomerAddress { get; set; }
        public required string CustomerContactNumber { get; set; }
        public required string BillCode { get; set; }
        public int BillYear { get; set; }
        public int BillMonth { get; set; }
        public BillingMethodType BillingMethod { get; set; }
        public ReceiveItemType ReceiveItem { get; set; }
        public bool IsSplitByReceiveItem { get; set; }
        public required string IndustCategory { get; set; }
        public string? EelecBelong { get; set; }
        public DateTime PaymentDeadline { get; set; }
        public int TotalPowerUse { get; set; }
        public DateTime SettleTime { get; set; }
        public DateTime ModifyTime { get; set; }
        public Guid Modifier { get; set; }
        public List<ExcelColumnCustomerContractBillItem> Items { get; set; } = new List<ExcelColumnCustomerContractBillItem>();
    }

    public class ExcelColumnCustomerContractBillItem
    {
        public Guid? CustomerContractBillItemId { get; set; }
        public Guid CustomerContractBillId { get; set; }
        public int SortOrder { get; set; }
        public required string ItemName { get; set; }
        public required string ProjectId { get; set; }
        public required string Unit { get; set; }
        public decimal? UnitPrice { get; set; }
        public int? Quantity { get; set; }
        public int TotalAmount { get; set; }
        public string? Note { get; set; }
        public Guid? Modifier { get; set; }
        public bool IsExtra { get; set; } = false;
    }

    public enum BillingMethodType
    {
        None,                           // 沒有選擇任何資料
        GroupByBelong,                  // 依照電號所屬合併開立
        GroupByBelongAndSeparateByPlace // 依照電號所屬合併開立 (區分案場)
    }

    public enum ReceiveItemType
    {
        None,                              // 沒有選擇
        OnlyElec,                          // 電費
        ElecRedistribute,                  // 電費+轉供費
        ElecRedistributeCertAudit,         // 電費+轉供費+憑證審查費
        ElecRedistributeCertAuditCertServ, // 電費+轉供費+憑證審查費+憑證服務費
        All                                // 全部
    }

}
