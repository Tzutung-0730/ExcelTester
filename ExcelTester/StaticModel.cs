using Bogus;
using ExcelTester.Classes;

namespace ExcelTester
{
    internal static class StaticModel
    {
        public static IEnumerable<ExcelColumnReport> GenerateFakeData()
        {
            var faker = new Faker<ExcelColumnReport>()
                .RuleFor(o => o.RowID, f => f.IndexFaker + 1)
                .RuleFor(o => o.ADYear, f => f.Date.Past(50).Year)
                .RuleFor(o => o.BuildCity, f => f.PickRandom(new[] { "台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市", "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣", "台東縣", "澎湖縣", "金門縣", "連江縣" }))
                .RuleFor(o => o.GovName, f => f.Name.FullName())
                .RuleFor(o => o.BuildName, f => f.Company.CompanyName())
                .RuleFor(o => o.BuildType, f => f.Commerce.Department())
                .RuleFor(o => o.BuildYear, f => f.Date.Past(100).Year)
                .RuleFor(o => o.HouseAge, (f, o) => DateTime.Now.Year - o.BuildYear)
                .RuleFor(o => o.FloorArea, f => f.Random.Decimal(50, 500))
                .RuleFor(o => o.PrelimBERS, f => f.Random.AlphaNumeric(10))
                .RuleFor(o => o.FormalBERS, f => f.Random.AlphaNumeric(10))
                .RuleFor(o => o.Tungsten, f => f.Random.Int(0, 100))
                .RuleFor(o => o.Fluorescent, f => f.Random.Int(0, 100))
                .RuleFor(o => o.Led, f => f.Random.Int(0, 100))
                .RuleFor(o => o.ElevType, f => f.Commerce.ProductName())
                .RuleFor(o => o.ElevCount, f => f.Random.Int(0, 10))
                .RuleFor(o => o.WHType, f => f.Commerce.ProductName())
                .RuleFor(o => o.Improve, f => f.Lorem.Sentence())
                .RuleFor(o => o.ImproveYear, f => f.Date.Past(20).Year.ToString())
                .RuleFor(o => o.SolarRoofKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.SolarGroundKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.TerritorialWatersKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.SolarMergeKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.SolarKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.HWindmillKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.VWindmillKW, f => f.Random.Decimal(0, 100))
                .RuleFor(o => o.WindKW, f => f.Random.Decimal(0, 100));

            List<ExcelColumnReport> reports = faker.Generate(100);

            return reports; 
        }
    }
}
