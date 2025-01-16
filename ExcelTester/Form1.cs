using ExcelTester.Classes;

namespace ExcelTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            var report = new Report();
            await report.GeneralReport("Templates/Report_Template.xlsx", "Report_Result.xlsx");
            MessageBox.Show("產生報表完成");
        }
    }
}
