namespace ExcelTester.Classes
{
    public class ExcelColumnTaipowerBill
    {
        public Guid TaipowerBillId { get; set; }
        public int BillYear { get; set; }
        public int BillMonth { get; set; }
        public DateTime PaymentDeadline { get; set; }
        public DateTime SettleTime { get; set; }
        public Guid Modifier { get; set; }
        public List<ExcelColumnTaipowerBillItem> Items { get; set; } = new List<ExcelColumnTaipowerBillItem>();
    }

    public class ExcelColumnTaipowerBillItem
    {
        public Guid TaipowerBillItemId { get; set; }
        public Guid TaipowerBillId { get; set; }
        public string? ProjectId { get; set; }
        public Guid? SupplierPlaceId { get; set; }
        public string? SupplierPlaceName { get; set; }
        public Guid CustomerId { get; set; }
        public required string CustomerName { get; set; }
        public int TotalAmount { get; set; }
        public Guid? Modifier { get; set; }
        public bool IsExtra { get; set; } = false;
    }
}
