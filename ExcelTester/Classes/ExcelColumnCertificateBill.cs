namespace ExcelTester.Classes
{
    public class ExcelColumnCertificateBill
    {
        public Guid CertificateBillId { get; set; }
        public int BillYear { get; set; }
        public int BillMonth { get; set; }
        public DateTime PaymentDeadline { get; set; }
        public DateTime SettleTime { get; set; }
        public DateTime ModifyTime { get; set; }
        public Guid Modifier { get; set; }
        public List<ExcelColumnCertificateBillItem> Items { get; set; } = new List<ExcelColumnCertificateBillItem>();
    }

    public class ExcelColumnCertificateBillItem
    {
        public Guid CertificateBillItemId { get; set; }
        public Guid CertificateBillId { get; set; }
        public string? ProjectId { get; set; }
        public Guid? SupplierPlaceId { get; set; }
        public required string SupplierPlaceName { get; set; }
        public int CertificateCount { get; set; }
        public int CertificateFee { get; set; }
        public int ServiceFee { get; set; }
        public DateTime? ModifyTime { get; set; }
        public Guid? Modifier { get; set; }
        public bool IsExtra { get; set; } = false;
    }
}
