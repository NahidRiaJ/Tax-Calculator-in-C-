namespace Tax_Calculator
{
    partial class texForm
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
            label1 = new Label();
            label2 = new Label();
            itemPrice = new TextBox();
            taxRate = new TextBox();
            addItem = new Button();
            totalCost = new Button();
            label3 = new Label();
            label4 = new Label();
            label5 = new Label();
            subTotal = new TextBox();
            taxAmount = new TextBox();
            total = new TextBox();
            dgvTaxCalculator = new DataGridView();
            label6 = new Label();
            btnExportToPDF = new Button();
            btnExportToExcel = new Button();
            ((System.ComponentModel.ISupportInitialize)dgvTaxCalculator).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(43, 33);
            label1.Name = "label1";
            label1.Size = new Size(75, 20);
            label1.TabIndex = 0;
            label1.Text = "Item Price";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(43, 84);
            label2.Name = "label2";
            label2.Size = new Size(64, 20);
            label2.TabIndex = 1;
            label2.Text = "Tax Rate";
            // 
            // itemPrice
            // 
            itemPrice.Location = new Point(147, 33);
            itemPrice.Multiline = true;
            itemPrice.Name = "itemPrice";
            itemPrice.Size = new Size(152, 34);
            itemPrice.TabIndex = 2;
            // 
            // taxRate
            // 
            taxRate.Location = new Point(147, 84);
            taxRate.Multiline = true;
            taxRate.Name = "taxRate";
            taxRate.Size = new Size(152, 34);
            taxRate.TabIndex = 3;
            // 
            // addItem
            // 
            addItem.BackColor = SystemColors.ActiveBorder;
            addItem.Location = new Point(43, 158);
            addItem.Name = "addItem";
            addItem.Size = new Size(256, 50);
            addItem.TabIndex = 4;
            addItem.Text = "Add Item";
            addItem.UseVisualStyleBackColor = false;
            addItem.Click += addItem_Click;
            // 
            // totalCost
            // 
            totalCost.BackColor = SystemColors.ActiveBorder;
            totalCost.Location = new Point(43, 226);
            totalCost.Name = "totalCost";
            totalCost.Size = new Size(256, 40);
            totalCost.TabIndex = 5;
            totalCost.Text = "Calculate Total Cost";
            totalCost.UseVisualStyleBackColor = false;
            totalCost.Click += totalCost_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label3.Location = new Point(43, 284);
            label3.Name = "label3";
            label3.Size = new Size(70, 20);
            label3.TabIndex = 6;
            label3.Text = "SubTotal";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label4.Location = new Point(43, 332);
            label4.Name = "label4";
            label4.Size = new Size(95, 20);
            label4.TabIndex = 7;
            label4.Text = "Tax Amount";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label5.Location = new Point(43, 384);
            label5.Name = "label5";
            label5.Size = new Size(44, 20);
            label5.TabIndex = 8;
            label5.Text = "Total";
            // 
            // subTotal
            // 
            subTotal.Location = new Point(147, 281);
            subTotal.Multiline = true;
            subTotal.Name = "subTotal";
            subTotal.Size = new Size(125, 34);
            subTotal.TabIndex = 9;
            // 
            // taxAmount
            // 
            taxAmount.Location = new Point(147, 332);
            taxAmount.Multiline = true;
            taxAmount.Name = "taxAmount";
            taxAmount.Size = new Size(125, 34);
            taxAmount.TabIndex = 10;
            // 
            // total
            // 
            total.Location = new Point(147, 381);
            total.Multiline = true;
            total.Name = "total";
            total.Size = new Size(125, 34);
            total.TabIndex = 11;
            // 
            // dgvTaxCalculator
            // 
            dgvTaxCalculator.AllowUserToAddRows = false;
            dgvTaxCalculator.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvTaxCalculator.Location = new Point(449, 68);
            dgvTaxCalculator.Name = "dgvTaxCalculator";
            dgvTaxCalculator.RowHeadersWidth = 51;
            dgvTaxCalculator.Size = new Size(283, 369);
            dgvTaxCalculator.TabIndex = 12;
            dgvTaxCalculator.CellClick += dgvTaxCalculator_CellClick;
            dgvTaxCalculator.CellFormatting += dgvTaxCalculator_CellFormatting;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 16.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label6.Location = new Point(496, 9);
            label6.Name = "label6";
            label6.Size = new Size(166, 38);
            label6.TabIndex = 13;
            label6.Text = "List of Items";
            // 
            // btnExportToPDF
            // 
            btnExportToPDF.Location = new Point(414, 492);
            btnExportToPDF.Name = "btnExportToPDF";
            btnExportToPDF.Size = new Size(116, 29);
            btnExportToPDF.TabIndex = 15;
            btnExportToPDF.Text = "Export to PDF";
            btnExportToPDF.UseVisualStyleBackColor = true;
            btnExportToPDF.Click += btnExportToPDF_Click;
            // 
            // btnExportToExcel
            // 
            btnExportToExcel.Location = new Point(626, 492);
            btnExportToExcel.Name = "btnExportToExcel";
            btnExportToExcel.Size = new Size(132, 29);
            btnExportToExcel.TabIndex = 16;
            btnExportToExcel.Text = "Export to Excel";
            btnExportToExcel.UseVisualStyleBackColor = true;
            btnExportToExcel.Click += btnExportToExcel_Click;
            // 
            // texForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(812, 627);
            Controls.Add(btnExportToExcel);
            Controls.Add(btnExportToPDF);
            Controls.Add(label6);
            Controls.Add(dgvTaxCalculator);
            Controls.Add(total);
            Controls.Add(taxAmount);
            Controls.Add(subTotal);
            Controls.Add(label5);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(totalCost);
            Controls.Add(addItem);
            Controls.Add(taxRate);
            Controls.Add(itemPrice);
            Controls.Add(label2);
            Controls.Add(label1);
            ForeColor = Color.Black;
            Name = "texForm";
            Text = "MainWindow";
            Load += texForm_Load;
            ((System.ComponentModel.ISupportInitialize)dgvTaxCalculator).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Label label2;
        private TextBox itemPrice;
        private TextBox taxRate;
        private Button addItem;
        private Button totalCost;
        private Label label3;
        private Label label4;
        private Label label5;
        private TextBox subTotal;
        private TextBox taxAmount;
        private TextBox total;
        private DataGridView dgvTaxCalculator;
        private Label label6;
        private Button btnExportToPDF;
        private Button btnExportToExcel;
    }
}
