using ExcelTester.Classes;

namespace ExcelTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void TaipowerBill_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelTaipowerBill.GenerateFakeData();

                TaipowerBill.GenerateTaipowerBill(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }
    }
}
