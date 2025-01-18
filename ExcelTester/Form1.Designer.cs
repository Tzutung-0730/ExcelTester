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
            btnCustomerContractBill1 = new Button();
            btnCustomerContractBill2 = new Button();
            SuspendLayout();
            // 
            // btnTaipowerBill
            // 
            btnTaipowerBill.Location = new Point(540, 94);
            btnTaipowerBill.Margin = new Padding(4);
            btnTaipowerBill.Name = "btnTaipowerBill";
            btnTaipowerBill.Size = new Size(550, 80);
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
            btnCertificateBill.Location = new Point(540, 210);
            btnCertificateBill.Margin = new Padding(4);
            btnCertificateBill.Name = "btnCertificateBill";
            btnCertificateBill.Size = new Size(550, 78);
            btnCertificateBill.TabIndex = 1;
            btnCertificateBill.Text = "憑證中心憑證費用_請付款單 匯出";
            btnCertificateBill.UseVisualStyleBackColor = true;
            btnCertificateBill.Click += btnCertificateBill_Click;
            // 
            // btnSupplierContractBill1
            // 
            btnSupplierContractBill1.Location = new Point(540, 334);
            btnSupplierContractBill1.Margin = new Padding(4);
            btnSupplierContractBill1.Name = "btnSupplierContractBill1";
            btnSupplierContractBill1.Size = new Size(550, 78);
            btnSupplierContractBill1.TabIndex = 2;
            btnSupplierContractBill1.Text = "台汽綠能_電業公司_請付款單 匯出";
            btnSupplierContractBill1.UseVisualStyleBackColor = true;
            btnSupplierContractBill1.Click += btnSupplierContractBill1_Click;
            // 
            // btnSupplierContractBill2
            // 
            btnSupplierContractBill2.Location = new Point(540, 460);
            btnSupplierContractBill2.Margin = new Padding(4);
            btnSupplierContractBill2.Name = "btnSupplierContractBill2";
            btnSupplierContractBill2.Size = new Size(550, 78);
            btnSupplierContractBill2.TabIndex = 3;
            btnSupplierContractBill2.Text = "台汽綠能_電業公司_付款通知 匯出";
            btnSupplierContractBill2.UseVisualStyleBackColor = true;
            btnSupplierContractBill2.Click += btnSupplierContractBill2_Click;
            // 
            // btnCustomerContractBill1
            // 
            btnCustomerContractBill1.Location = new Point(540, 590);
            btnCustomerContractBill1.Margin = new Padding(4);
            btnCustomerContractBill1.Name = "btnCustomerContractBill1";
            btnCustomerContractBill1.Size = new Size(550, 78);
            btnCustomerContractBill1.TabIndex = 4;
            btnCustomerContractBill1.Text = "台汽綠能_客戶_統一發票資料單 匯出";
            btnCustomerContractBill1.UseVisualStyleBackColor = true;
            btnCustomerContractBill1.Click += btnCustomerContractBill1_Click;
            // 
            // btnCustomerContractBill2
            // 
            btnCustomerContractBill2.Location = new Point(540, 715);
            btnCustomerContractBill2.Margin = new Padding(4);
            btnCustomerContractBill2.Name = "btnCustomerContractBill2";
            btnCustomerContractBill2.Size = new Size(550, 78);
            btnCustomerContractBill2.TabIndex = 5;
            btnCustomerContractBill2.Text = "台汽綠能_客戶_繳費通知 匯出";
            btnCustomerContractBill2.UseVisualStyleBackColor = true;
            btnCustomerContractBill2.Click += btnCustomerContractBill2_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(18F, 38F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1698, 876);
            Controls.Add(btnCustomerContractBill2);
            Controls.Add(btnCustomerContractBill1);
            Controls.Add(btnSupplierContractBill2);
            Controls.Add(btnSupplierContractBill1);
            Controls.Add(btnCertificateBill);
            Controls.Add(btnTaipowerBill);
            Margin = new Padding(4);
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
        private Button btnCustomerContractBill1;
        private Button btnCustomerContractBill2;
    }
}
