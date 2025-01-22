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
                var data = StaticModelTaipowerBill.GetData();
                TaipowerBill.WriteExcel(fileName, newFileName, data);
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
                var data = StaticModelCertificateBill.GetData();
                CertificateBill.WriteExcel(fileName, newFileName, data);
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
                var data = StaticModelSupplierContractBill.GetData();
                SupplierContractBill.WriteExcel_Invoice(fileName, newFileName, data);
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
                var data = StaticModelSupplierContractBill.GetData();
                SupplierContractBill.WriteExcel_PaymentNotice(fileName, newFileName, data);
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
                var data = StaticModelCustomerContractBill.GetData();
                CustomerContractBill.WriteExcel_Invoice(fileName, newFileName, data);
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
                var data = StaticModelCustomerContractBill.GetData();
                CustomerContractBill.WriteExcel_PaymentNotice(fileName, newFileName, data);
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
                var data = StaticModelSupplierContractBill.GetData();
                SupplierContractBillWord.WriteWord(fileName, newFileName, data);
            }

            MessageBox.Show("Done");
        }
    }
}
