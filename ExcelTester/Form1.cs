using ExcelTester.Classes;

namespace ExcelTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnTaipowerBill_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelTaipowerBill.GenerateFakeData();
                TaipowerBill.WriteExcel(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }

        private void btnCertificateBill_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelCertificateBill.GenerateFakeData();
                CertificateBill.WriteExcel(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }

        private void btnSupplierContractBill1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelSupplierContractBill.GenerateFakeData();
                SupplierContractBill.WriteExcel_Invoice(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }

        private void btnSupplierContractBill2_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelSupplierContractBill.GenerateFakeData();
                SupplierContractBill.WriteExcel_PaymentNotice(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }

        private void btnCustomerContractBill1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelCustomerContractBill.GenerateFakeData();
                CustomerContractBill.WriteExcel_Invoice(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }

        private void btnCustomerContractBill2_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Excel ¿… (*.xlsx)|*.xlsx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelCustomerContractBill.GenerateFakeData();
                CustomerContractBill.WriteExcel_PaymentNotice(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }

        private void btnSupplierContractBill3_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "Word ¿… (*.docx)|*.docx";

            var openResult = openFileDialog1.ShowDialog();
            if (openResult == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                var newFileName = $"{Path.GetDirectoryName(fileName)}/{Path.GetFileNameWithoutExtension(fileName)}1{Path.GetExtension(fileName)}";
                var fakeData = StaticModelSupplierContractBill.GenerateFakeData();
                SupplierContractBillWord.WriteWord(fileName, newFileName, fakeData);
            }

            MessageBox.Show("Done");
        }
    }
}
