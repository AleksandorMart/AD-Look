namespace ADlook
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnFreeSearch = new Button();
            dgvResults = new DataGridView();
            txt1 = new TextBox();
            cmbQueries = new ComboBox();
            pnlFreeSearch = new Panel();
            cmbFreeSearch = new ComboBox();
            txt2 = new TextBox();
            label2 = new Label();
            label1 = new Label();
            lblStatus = new Label();
            btnExport = new Button();
            pnlOUSearch = new Panel();
            txt3 = new TextBox();
            label3 = new Label();
            pnlTimeSearch = new Panel();
            dtpTimeSearcher = new DateTimePicker();
            label4 = new Label();
            pnlOSSearch = new Panel();
            label7 = new Label();
            label6 = new Label();
            txt4 = new TextBox();
            label5 = new Label();
            ((System.ComponentModel.ISupportInitialize)dgvResults).BeginInit();
            pnlFreeSearch.SuspendLayout();
            pnlOUSearch.SuspendLayout();
            pnlTimeSearch.SuspendLayout();
            pnlOSSearch.SuspendLayout();
            SuspendLayout();
            // 
            // btnFreeSearch
            // 
            btnFreeSearch.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnFreeSearch.Location = new Point(639, 377);
            btnFreeSearch.Name = "btnFreeSearch";
            btnFreeSearch.Size = new Size(149, 23);
            btnFreeSearch.TabIndex = 0;
            btnFreeSearch.Text = "Поиск";
            btnFreeSearch.UseVisualStyleBackColor = true;
            btnFreeSearch.Click += btnFreeSearch_Click;
            // 
            // dgvResults
            // 
            dgvResults.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgvResults.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvResults.Location = new Point(12, 41);
            dgvResults.Name = "dgvResults";
            dgvResults.RowHeadersWidth = 62;
            dgvResults.RowTemplate.Height = 25;
            dgvResults.Size = new Size(621, 397);
            dgvResults.TabIndex = 1;
            dgvResults.CellDoubleClick += dgvResults_CellDoubleClick;
            // 
            // txt1
            // 
            txt1.Location = new Point(3, 47);
            txt1.Name = "txt1";
            txt1.Size = new Size(132, 23);
            txt1.TabIndex = 2;
            // 
            // cmbQueries
            // 
            cmbQueries.FormattingEnabled = true;
            cmbQueries.Location = new Point(12, 12);
            cmbQueries.Name = "cmbQueries";
            cmbQueries.Size = new Size(149, 23);
            cmbQueries.TabIndex = 3;
            cmbQueries.SelectedIndexChanged += cmbQueries_SelectedIndexChanged;
            // 
            // pnlFreeSearch
            // 
            pnlFreeSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pnlFreeSearch.BorderStyle = BorderStyle.FixedSingle;
            pnlFreeSearch.Controls.Add(cmbFreeSearch);
            pnlFreeSearch.Controls.Add(txt2);
            pnlFreeSearch.Controls.Add(label2);
            pnlFreeSearch.Controls.Add(label1);
            pnlFreeSearch.Controls.Add(txt1);
            pnlFreeSearch.Location = new Point(639, 41);
            pnlFreeSearch.Name = "pnlFreeSearch";
            pnlFreeSearch.Size = new Size(149, 123);
            pnlFreeSearch.TabIndex = 4;
            // 
            // cmbFreeSearch
            // 
            cmbFreeSearch.FormattingEnabled = true;
            cmbFreeSearch.Location = new Point(3, 3);
            cmbFreeSearch.Name = "cmbFreeSearch";
            cmbFreeSearch.Size = new Size(132, 23);
            cmbFreeSearch.TabIndex = 6;
            cmbFreeSearch.SelectedIndexChanged += cmbFreeSearch_SelectedIndexChanged;
            // 
            // txt2
            // 
            txt2.Location = new Point(3, 91);
            txt2.Name = "txt2";
            txt2.Size = new Size(132, 23);
            txt2.TabIndex = 5;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(3, 29);
            label2.Name = "label2";
            label2.Size = new Size(34, 15);
            label2.TabIndex = 4;
            label2.Text = "ФИО";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(3, 73);
            label1.Name = "label1";
            label1.Size = new Size(41, 15);
            label1.TabIndex = 3;
            label1.Text = "Логин";
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(167, 20);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(114, 15);
            lblStatus.TabIndex = 5;
            lblStatus.Text = "Найдено: 0 записей";
            // 
            // btnExport
            // 
            btnExport.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnExport.Location = new Point(639, 406);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(149, 31);
            btnExport.TabIndex = 6;
            btnExport.Text = "Экспорт в Excel";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnExport_Click;
            // 
            // pnlOUSearch
            // 
            pnlOUSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pnlOUSearch.BorderStyle = BorderStyle.FixedSingle;
            pnlOUSearch.Controls.Add(txt3);
            pnlOUSearch.Controls.Add(label3);
            pnlOUSearch.Location = new Point(639, 170);
            pnlOUSearch.Name = "pnlOUSearch";
            pnlOUSearch.Size = new Size(149, 53);
            pnlOUSearch.TabIndex = 7;
            // 
            // txt3
            // 
            txt3.Location = new Point(3, 20);
            txt3.Name = "txt3";
            txt3.Size = new Size(132, 23);
            txt3.TabIndex = 1;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(3, 2);
            label3.Name = "label3";
            label3.Size = new Size(40, 15);
            label3.TabIndex = 0;
            label3.Text = "Отдел";
            // 
            // pnlTimeSearch
            // 
            pnlTimeSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pnlTimeSearch.BorderStyle = BorderStyle.FixedSingle;
            pnlTimeSearch.Controls.Add(dtpTimeSearcher);
            pnlTimeSearch.Controls.Add(label4);
            pnlTimeSearch.Location = new Point(639, 229);
            pnlTimeSearch.Name = "pnlTimeSearch";
            pnlTimeSearch.Size = new Size(149, 55);
            pnlTimeSearch.TabIndex = 8;
            // 
            // dtpTimeSearcher
            // 
            dtpTimeSearcher.Location = new Point(3, 20);
            dtpTimeSearcher.Name = "dtpTimeSearcher";
            dtpTimeSearcher.Size = new Size(132, 23);
            dtpTimeSearcher.TabIndex = 1;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(3, 2);
            label4.Name = "label4";
            label4.Size = new Size(89, 15);
            label4.TabIndex = 0;
            label4.Text = "Время отсечки";
            // 
            // pnlOSSearch
            // 
            pnlOSSearch.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pnlOSSearch.BorderStyle = BorderStyle.FixedSingle;
            pnlOSSearch.Controls.Add(label7);
            pnlOSSearch.Controls.Add(label6);
            pnlOSSearch.Controls.Add(txt4);
            pnlOSSearch.Controls.Add(label5);
            pnlOSSearch.Location = new Point(639, 290);
            pnlOSSearch.Name = "pnlOSSearch";
            pnlOSSearch.Size = new Size(149, 81);
            pnlOSSearch.TabIndex = 9;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(3, 32);
            label7.Name = "label7";
            label7.Size = new Size(52, 15);
            label7.TabIndex = 4;
            label7.Text = "запятую";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(3, 17);
            label6.Name = "label6";
            label6.Size = new Size(105, 15);
            label6.TabIndex = 3;
            label6.Text = "перчислять через";
            // 
            // txt4
            // 
            txt4.Location = new Point(3, 50);
            txt4.Name = "txt4";
            txt4.Size = new Size(132, 23);
            txt4.TabIndex = 1;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(3, 2);
            label5.Name = "label5";
            label5.Size = new Size(24, 15);
            label5.TabIndex = 0;
            label5.Text = "ОС";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 449);
            Controls.Add(pnlOSSearch);
            Controls.Add(pnlTimeSearch);
            Controls.Add(pnlOUSearch);
            Controls.Add(btnExport);
            Controls.Add(lblStatus);
            Controls.Add(pnlFreeSearch);
            Controls.Add(btnFreeSearch);
            Controls.Add(cmbQueries);
            Controls.Add(dgvResults);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "AD Look";
            ((System.ComponentModel.ISupportInitialize)dgvResults).EndInit();
            pnlFreeSearch.ResumeLayout(false);
            pnlFreeSearch.PerformLayout();
            pnlOUSearch.ResumeLayout(false);
            pnlOUSearch.PerformLayout();
            pnlTimeSearch.ResumeLayout(false);
            pnlTimeSearch.PerformLayout();
            pnlOSSearch.ResumeLayout(false);
            pnlOSSearch.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnFreeSearch;
        private DataGridView dgvResults;
        private TextBox txt1;
        private ComboBox cmbQueries;
        private Panel pnlFreeSearch;
        private Label label1;
        private TextBox txt2;
        private Label label2;
        private Label lblStatus;
        private ComboBox cmbFreeSearch;
        private Button btnExport;
        private Panel pnlOUSearch;
        private TextBox txt3;
        private Label label3;
        private Panel pnlTimeSearch;
        private Label label4;
        private DateTimePicker dtpTimeSearcher;
        private Panel pnlOSSearch;
        private TextBox txt4;
        private Label label5;
        private Label label7;
        private Label label6;
    }
}