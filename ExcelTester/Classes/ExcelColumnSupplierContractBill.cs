namespace ExcelTester.Classes
{
    public class ExcelColumnSupplierContractBill
    {
        public Guid SupplierContractBillId { get; set; }
        public Guid SupplierId { get; set; }
        public required string SupplierName { get; set; }
        public Guid SupplierContractId { get; set; }
        public required string SupplierTaxIDNumber { get; set; }
        public required string SupplierAddress { get; set; }
        public required string SupplierContactNumber { get; set; }
        public required string BillCode { get; set; }
        public int BillYear { get; set; }
        public int BillMonth { get; set; }
        public DateTime PaymentDeadline { get; set; }
        public required string PaymentMethod { get; set; }
        public string? ReceiveAccount { get; set; }
        public string? ReceiveName { get; set; }
        public string? ReceiveBankName { get; set; }
        public DateTime SettleTime { get; set; }
        public DateTime ModifyTime { get; set; }
        public Guid Modifier { get; set; }
        public List<ExcelColumnSupplierContractBillItem> Items { get; set; } = new List<ExcelColumnSupplierContractBillItem>();
    }

    public class ExcelColumnSupplierContractBillItem
    {
        public Guid SupplierContractBillItemId { get; set; }
        public Guid SupplierContractBillId { get; set; }
        public int SortOrder { get; set; }
        public required string ItemName { get; set; }
        public Guid SupplierPlaceId { get; set; }
        public required string SupplierPlaceName { get; set; }
        public required string SupplierElectricityNumber { get; set; }
        public required string ProjectId { get; set; }
        public required string Unit { get; set; }
        public decimal? UnitPrice { get; set; }
        public int? Quantity { get; set; }
        public int TotalAmount { get; set; }
        public string? Note { get; set; }
        public DateTime? ModifyTime { get; set; }
        public Guid? Modifier { get; set; }
        public bool IsExtra { get; set; }
    }
}
