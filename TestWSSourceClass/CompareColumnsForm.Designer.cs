namespace Converter
{
    partial class CompareColumnsForm
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
            this.button1 = new System.Windows.Forms.Button();
            this.addColumnButton = new System.Windows.Forms.Button();
            this.UnUsedColumnListBox = new System.Windows.Forms.ListBox();
            this.UnUsedSourceSolumnLanbel = new System.Windows.Forms.Label();
            this.ColumnExampleListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SourceDescrLabel = new System.Windows.Forms.Label();
            this.TemplateDescrLabel = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button1.BackColor = System.Drawing.SystemColors.Control;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(490, 549);
            this.button1.Margin = new System.Windows.Forms.Padding(10);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(172, 29);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // addColumnButton
            // 
            this.addColumnButton.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.addColumnButton.BackColor = System.Drawing.SystemColors.Control;
            this.addColumnButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.addColumnButton.Location = new System.Drawing.Point(490, 3);
            this.addColumnButton.Name = "addColumnButton";
            this.addColumnButton.Size = new System.Drawing.Size(25, 31);
            this.addColumnButton.TabIndex = 2;
            this.addColumnButton.Text = "+";
            this.addColumnButton.UseVisualStyleBackColor = false;
            this.addColumnButton.Click += new System.EventHandler(this.addColumnButton_Click);
            // 
            // UnUsedColumnListBox
            // 
            this.UnUsedColumnListBox.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.UnUsedColumnListBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.UnUsedColumnListBox.FormattingEnabled = true;
            this.UnUsedColumnListBox.Location = new System.Drawing.Point(490, 70);
            this.UnUsedColumnListBox.Name = "UnUsedColumnListBox";
            this.UnUsedColumnListBox.Size = new System.Drawing.Size(172, 210);
            this.UnUsedColumnListBox.TabIndex = 4;
            this.UnUsedColumnListBox.SelectedIndexChanged += new System.EventHandler(this.UnSuedColumnListBox_SelectedIndexChanged);
            this.UnUsedColumnListBox.DoubleClick += new System.EventHandler(this.UnUsedColumnListBox_DoubleClick);
            // 
            // UnUsedSourceSolumnLanbel
            // 
            this.UnUsedSourceSolumnLanbel.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.UnUsedSourceSolumnLanbel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.UnUsedSourceSolumnLanbel.Location = new System.Drawing.Point(487, 37);
            this.UnUsedSourceSolumnLanbel.Name = "UnUsedSourceSolumnLanbel";
            this.UnUsedSourceSolumnLanbel.Size = new System.Drawing.Size(172, 30);
            this.UnUsedSourceSolumnLanbel.TabIndex = 5;
            this.UnUsedSourceSolumnLanbel.Text = "Нераспределенные столбцы выгрузки";
            // 
            // ColumnExampleListBox
            // 
            this.ColumnExampleListBox.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.ColumnExampleListBox.FormattingEnabled = true;
            this.ColumnExampleListBox.Location = new System.Drawing.Point(490, 325);
            this.ColumnExampleListBox.Name = "ColumnExampleListBox";
            this.ColumnExampleListBox.Size = new System.Drawing.Size(172, 199);
            this.ColumnExampleListBox.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label1.Location = new System.Drawing.Point(487, 295);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(172, 27);
            this.label1.TabIndex = 7;
            this.label1.Text = "Первые 15 значений по выбранному столбцу";
            // 
            // SourceDescrLabel
            // 
            this.SourceDescrLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SourceDescrLabel.AutoSize = true;
            this.SourceDescrLabel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.SourceDescrLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.SourceDescrLabel.Location = new System.Drawing.Point(245, 1);
            this.SourceDescrLabel.Name = "SourceDescrLabel";
            this.SourceDescrLabel.Size = new System.Drawing.Size(225, 26);
            this.SourceDescrLabel.TabIndex = 1;
            this.SourceDescrLabel.Tag = "Descr";
            this.SourceDescrLabel.Text = "Столбцы из выгрузки";
            this.SourceDescrLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // TemplateDescrLabel
            // 
            this.TemplateDescrLabel.AutoSize = true;
            this.TemplateDescrLabel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.TemplateDescrLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TemplateDescrLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TemplateDescrLabel.Location = new System.Drawing.Point(4, 1);
            this.TemplateDescrLabel.Name = "TemplateDescrLabel";
            this.TemplateDescrLabel.Size = new System.Drawing.Size(234, 26);
            this.TemplateDescrLabel.TabIndex = 0;
            this.TemplateDescrLabel.Tag = "Descr";
            this.TemplateDescrLabel.Text = "Столбцы по шаблону";
            this.TemplateDescrLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.AutoScroll = true;
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.BackColor = System.Drawing.SystemColors.Control;
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 240F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 231F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.TemplateDescrLabel, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.SourceDescrLabel, 1, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(12, 4);
            this.tableLayoutPanel1.MaximumSize = new System.Drawing.Size(979, 574);
            this.tableLayoutPanel1.MinimumSize = new System.Drawing.Size(472, 20);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(474, 28);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // CompareColumnsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(674, 597);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ColumnExampleListBox);
            this.Controls.Add(this.UnUsedSourceSolumnLanbel);
            this.Controls.Add(this.UnUsedColumnListBox);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.addColumnButton);
            this.Controls.Add(this.button1);
            this.MaximumSize = new System.Drawing.Size(1200, 636);
            this.MinimumSize = new System.Drawing.Size(690, 636);
            this.Name = "CompareColumnsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Сопоставьте колонки";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.CompareColumnsForm_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void CompareColumnsForm_Load(object sender, System.EventArgs e)
        {
            //throw new System.NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.Button addColumnButton;
        private System.Windows.Forms.ListBox UnUsedColumnListBox;
        private System.Windows.Forms.Label UnUsedSourceSolumnLanbel;
        private System.Windows.Forms.ListBox ColumnExampleListBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label SourceDescrLabel;
        private System.Windows.Forms.Label TemplateDescrLabel;
        public System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        public System.Windows.Forms.Button button1;


    }
}