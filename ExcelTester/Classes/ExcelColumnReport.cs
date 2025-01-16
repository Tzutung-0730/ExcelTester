using ExcelTester.Attributes;

namespace ExcelTester.Classes
{
    public class ExcelColumnReport
    {
        [ExcelColumnMapping("A")]
        public int RowID { get; set; }

        [ExcelColumnMapping("B")]
        public int ADYear { get; set; }

        [ExcelColumnMapping("C")]
        public string BuildCity { get; set; }

        [ExcelColumnMapping("D")]
        public string GovName { get; set; }

        [ExcelColumnMapping("E")]
        public string BuildName { get; set; }

        [ExcelColumnMapping("F")]
        public string BuildType { get; set; }

        [ExcelColumnMapping("G")]
        public int BuildYear { get; set; }

        [ExcelColumnMapping("H")]
        public int HouseAge { get; set; }

        [ExcelColumnMapping("I")]
        public decimal FloorArea { get; set; }

        [ExcelColumnMapping("J")]
        public string PrelimBERS { get; set; }

        [ExcelColumnMapping("K")]
        public string FormalBERS { get; set; }

        [ExcelColumnMapping("L")]
        public int Tungsten { get; set; }

        [ExcelColumnMapping("M")]
        public int Fluorescent { get; set; }

        [ExcelColumnMapping("N")]
        public int Led { get; set; }

        [ExcelColumnMapping("O")]
        public string ElevType { get; set; }

        [ExcelColumnMapping("P")]
        public int ElevCount { get; set; }

        [ExcelColumnMapping("Q")]
        public string WHType { get; set; }

        [ExcelColumnMapping("R")]
        public string Improve { get; set; }

        [ExcelColumnMapping("S")]
        public string ImproveYear { get; set; }

        [ExcelColumnMapping("T")]
        public decimal SolarRoofKW { get; set; }

        [ExcelColumnMapping("U")]
        public decimal SolarGroundKW { get; set; }

        [ExcelColumnMapping("V")]
        public decimal TerritorialWatersKW { get; set; }

        [ExcelColumnMapping("W")]
        public decimal SolarMergeKW { get; set; }

        [ExcelColumnMapping("X")]
        public decimal SolarKW { get; set; }

        [ExcelColumnMapping("Y")]
        public decimal HWindmillKW { get; set; }

        [ExcelColumnMapping("Z")]
        public decimal VWindmillKW { get; set; }

        [ExcelColumnMapping("AA")]
        public decimal WindKW { get; set; }
    }
}
