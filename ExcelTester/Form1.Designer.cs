namespace ExcelTester
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnTaipowerBill = new Button();
            openFileDialog1 = new OpenFileDialog();
            btnCertificateBill = new Button();
            btnSupplierContractBill1 = new Button();
            btnSupplierContractBill2 = new Button();
            SuspendLayout();
            // 
            // btnTaipowerBill
            // 
            btnTaipowerBill.Location = new Point(270, 47);
            btnTaipowerBill.Margin = new Padding(2);
            btnTaipowerBill.Name = "btnTaipowerBill";
            btnTaipowerBill.Size = new Size(275, 40);
            btnTaipowerBill.TabIndex = 0;
            btnTaipowerBill.Text = "台電轉供費用_請付款單 匯出";
            btnTaipowerBill.UseVisualStyleBackColor = true;
            btnTaipowerBill.Click += btnTaipowerBill_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnCertificateBill
            // 
            btnCertificateBill.Location = new Point(270, 105);
            btnCertificateBill.Margin = new Padding(2);
            btnCertificateBill.Name = "btnCertificateBill";
            btnCertificateBill.Size = new Size(275, 39);
            btnCertificateBill.TabIndex = 1;
            btnCertificateBill.Text = "憑證中心憑證費用_請付款單 匯出";
            btnCertificateBill.UseVisualStyleBackColor = true;
            btnCertificateBill.Click += btnCertificateBill_Click;
            // 
            // btnSupplierContractBill1
            // 
            btnSupplierContractBill1.Location = new Point(270, 167);
            btnSupplierContractBill1.Margin = new Padding(2);
            btnSupplierContractBill1.Name = "btnSupplierContractBill1";
            btnSupplierContractBill1.Size = new Size(275, 39);
            btnSupplierContractBill1.TabIndex = 2;
            btnSupplierContractBill1.Text = "台汽綠能_電業公司_請付款單 匯出";
            btnSupplierContractBill1.UseVisualStyleBackColor = true;
            btnSupplierContractBill1.Click += btnSupplierContractBill1_Click;
            // 
            // btnSupplierContractBill2
            // 
            btnSupplierContractBill2.Location = new Point(270, 230);
            btnSupplierContractBill2.Margin = new Padding(2);
            btnSupplierContractBill2.Name = "btnSupplierContractBill2";
            btnSupplierContractBill2.Size = new Size(275, 39);
            btnSupplierContractBill2.TabIndex = 3;
            btnSupplierContractBill2.Text = "台汽綠能_電業公司_付款通知 匯出";
            btnSupplierContractBill2.UseVisualStyleBackColor = true;
            btnSupplierContractBill2.Click += btnSupplierContractBill2_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(9F, 19F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(849, 438);
            Controls.Add(btnSupplierContractBill2);
            Controls.Add(btnSupplierContractBill1);
            Controls.Add(btnCertificateBill);
            Controls.Add(btnTaipowerBill);
            Margin = new Padding(2);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
        }

        #endregion

        private Button btnTaipowerBill;
        private OpenFileDialog openFileDialog1;
        private Button btnCertificateBill;
        private Button btnSupplierContractBill1;
        private Button btnSupplierContractBill2;
    }
}
