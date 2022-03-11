namespace OutputCashFlowExcel
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.Export_Excel_Tab = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.ProjectCodeTo_comboBox = new System.Windows.Forms.ComboBox();
            this.budgetCashFlowProjectBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.financial_Rpt_TestDataSet = new OutputCashFlowExcel.Financial_Rpt_TestDataSet();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.Date_comboBox = new System.Windows.Forms.ComboBox();
            this.datelistBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.financial_Rpt_TestDataSet2 = new OutputCashFlowExcel.Financial_Rpt_TestDataSet2();
            this.ProjectCodeFrom_comboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Export_button = new System.Windows.Forms.Button();
            this.budget_CashFlow_ProjectTableAdapter = new OutputCashFlowExcel.Financial_Rpt_TestDataSetTableAdapters.Budget_CashFlow_ProjectTableAdapter();
            this.financialRptTestDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.datelistTableAdapter = new OutputCashFlowExcel.Financial_Rpt_TestDataSet2TableAdapters.datelistTableAdapter();
            this.tabControl1.SuspendLayout();
            this.Export_Excel_Tab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.budgetCashFlowProjectBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.financial_Rpt_TestDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.datelistBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.financial_Rpt_TestDataSet2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.financialRptTestDataSetBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.Export_Excel_Tab);
            this.tabControl1.Location = new System.Drawing.Point(-2, 1);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(390, 248);
            this.tabControl1.TabIndex = 0;
            // 
            // Export_Excel_Tab
            // 
            this.Export_Excel_Tab.Controls.Add(this.label3);
            this.Export_Excel_Tab.Controls.Add(this.richTextBox1);
            this.Export_Excel_Tab.Controls.Add(this.ProjectCodeTo_comboBox);
            this.Export_Excel_Tab.Controls.Add(this.progressBar1);
            this.Export_Excel_Tab.Controls.Add(this.label2);
            this.Export_Excel_Tab.Controls.Add(this.Date_comboBox);
            this.Export_Excel_Tab.Controls.Add(this.ProjectCodeFrom_comboBox);
            this.Export_Excel_Tab.Controls.Add(this.label1);
            this.Export_Excel_Tab.Controls.Add(this.Export_button);
            this.Export_Excel_Tab.Location = new System.Drawing.Point(4, 21);
            this.Export_Excel_Tab.Name = "Export_Excel_Tab";
            this.Export_Excel_Tab.Padding = new System.Windows.Forms.Padding(3);
            this.Export_Excel_Tab.Size = new System.Drawing.Size(382, 223);
            this.Export_Excel_Tab.TabIndex = 0;
            this.Export_Excel_Tab.Text = "Export Excel";
            this.Export_Excel_Tab.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("PMingLiU", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(35, 108);
            this.label3.MinimumSize = new System.Drawing.Size(300, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(300, 30);
            this.label3.TabIndex = 8;
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(172, 129);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(163, 33);
            this.richTextBox1.TabIndex = 7;
            this.richTextBox1.Text = "";
            this.richTextBox1.Visible = false;
            // 
            // ProjectCodeTo_comboBox
            // 
            this.ProjectCodeTo_comboBox.DataSource = this.budgetCashFlowProjectBindingSource;
            this.ProjectCodeTo_comboBox.DisplayMember = "Budget_Project_No";
            this.ProjectCodeTo_comboBox.FormattingEnabled = true;
            this.ProjectCodeTo_comboBox.Location = new System.Drawing.Point(90, 66);
            this.ProjectCodeTo_comboBox.Name = "ProjectCodeTo_comboBox";
            this.ProjectCodeTo_comboBox.Size = new System.Drawing.Size(245, 20);
            this.ProjectCodeTo_comboBox.TabIndex = 6;
            this.ProjectCodeTo_comboBox.ValueMember = "Budget_Project_No";
            // 
            // budgetCashFlowProjectBindingSource
            // 
            this.budgetCashFlowProjectBindingSource.DataMember = "Budget_CashFlow_Project";
            this.budgetCashFlowProjectBindingSource.DataSource = this.financial_Rpt_TestDataSet;
            // 
            // financial_Rpt_TestDataSet
            // 
            this.financial_Rpt_TestDataSet.DataSetName = "Financial_Rpt_TestDataSet";
            this.financial_Rpt_TestDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(21, 168);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(314, 23);
            this.progressBar1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 96);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(26, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Date";
            // 
            // Date_comboBox
            // 
            this.Date_comboBox.DataSource = this.datelistBindingSource;
            this.Date_comboBox.DisplayMember = "YrEnd_Date";
            this.Date_comboBox.FormattingEnabled = true;
            this.Date_comboBox.Location = new System.Drawing.Point(90, 92);
            this.Date_comboBox.Name = "Date_comboBox";
            this.Date_comboBox.Size = new System.Drawing.Size(245, 20);
            this.Date_comboBox.TabIndex = 3;
            this.Date_comboBox.ValueMember = "YrEnd_Date";
            // 
            // datelistBindingSource
            // 
            this.datelistBindingSource.DataMember = "datelist";
            this.datelistBindingSource.DataSource = this.financial_Rpt_TestDataSet2;
            // 
            // financial_Rpt_TestDataSet2
            // 
            this.financial_Rpt_TestDataSet2.DataSetName = "Financial_Rpt_TestDataSet2";
            this.financial_Rpt_TestDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // ProjectCodeFrom_comboBox
            // 
            this.ProjectCodeFrom_comboBox.DataSource = this.budgetCashFlowProjectBindingSource;
            this.ProjectCodeFrom_comboBox.DisplayMember = "Budget_Project_No";
            this.ProjectCodeFrom_comboBox.FormattingEnabled = true;
            this.ProjectCodeFrom_comboBox.Location = new System.Drawing.Point(90, 37);
            this.ProjectCodeFrom_comboBox.Name = "ProjectCodeFrom_comboBox";
            this.ProjectCodeFrom_comboBox.Size = new System.Drawing.Size(245, 20);
            this.ProjectCodeFrom_comboBox.TabIndex = 2;
            this.ProjectCodeFrom_comboBox.ValueMember = "Budget_Project_No";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "Project Code";
            // 
            // Export_button
            // 
            this.Export_button.Location = new System.Drawing.Point(90, 129);
            this.Export_button.Name = "Export_button";
            this.Export_button.Size = new System.Drawing.Size(75, 23);
            this.Export_button.TabIndex = 0;
            this.Export_button.Text = "Export";
            this.Export_button.UseVisualStyleBackColor = true;
            this.Export_button.Click += new System.EventHandler(this.Export_button_Click);
            // 
            // budget_CashFlow_ProjectTableAdapter
            // 
            this.budget_CashFlow_ProjectTableAdapter.ClearBeforeFill = true;
            // 
            // financialRptTestDataSetBindingSource
            // 
            this.financialRptTestDataSetBindingSource.DataSource = this.financial_Rpt_TestDataSet;
            this.financialRptTestDataSetBindingSource.Position = 0;
            // 
            // datelistTableAdapter
            // 
            this.datelistTableAdapter.ClearBeforeFill = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(387, 248);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Exporting Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.Export_Excel_Tab.ResumeLayout(false);
            this.Export_Excel_Tab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.budgetCashFlowProjectBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.financial_Rpt_TestDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.datelistBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.financial_Rpt_TestDataSet2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.financialRptTestDataSetBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage Export_Excel_Tab;
        private System.Windows.Forms.Button Export_button;
        private System.Windows.Forms.ComboBox ProjectCodeFrom_comboBox;
        private System.Windows.Forms.Label label1;
        private Financial_Rpt_TestDataSet financial_Rpt_TestDataSet;
        private System.Windows.Forms.BindingSource budgetCashFlowProjectBindingSource;
        private OutputCashFlowExcel.Financial_Rpt_TestDataSetTableAdapters.Budget_CashFlow_ProjectTableAdapter budget_CashFlow_ProjectTableAdapter;
        private System.Windows.Forms.ComboBox Date_comboBox;
        private System.Windows.Forms.BindingSource financialRptTestDataSetBindingSource;
        private System.Windows.Forms.Label label2;
        private Financial_Rpt_TestDataSet2 financial_Rpt_TestDataSet2;
        private System.Windows.Forms.BindingSource datelistBindingSource;
        private OutputCashFlowExcel.Financial_Rpt_TestDataSet2TableAdapters.datelistTableAdapter datelistTableAdapter;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ComboBox ProjectCodeTo_comboBox;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label label3;
    }
}

