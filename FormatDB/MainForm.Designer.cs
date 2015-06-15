namespace Formater
{
    sealed partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.CatalogPathTextBox = new System.Windows.Forms.TextBox();
            this.OKTMOPathTextBox = new System.Windows.Forms.TextBox();
            this.SubjectSourcePathTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SelectCatalogFileButton = new System.Windows.Forms.Button();
            this.SelectSubjectSourceFileButton = new System.Windows.Forms.Button();
            this.SelectOKTMOFileButton = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.WarningLabel = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.OpenButton = new System.Windows.Forms.Button();
            this.workbookPathTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.StartButton = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label8 = new System.Windows.Forms.Label();
            this.OKTMOWorksheetCBox = new System.Windows.Forms.ComboBox();
            this.CatalogWorksheetCBox = new System.Windows.Forms.ComboBox();
            this.SubjectSourceWorksheetCBox = new System.Windows.Forms.ComboBox();
            this.SelectVGTFileButton = new System.Windows.Forms.Button();
            this.VGTPathTextBox = new System.Windows.Forms.TextBox();
            this.VGTWorksheetCBox = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CatalogPathTextBox
            // 
            this.CatalogPathTextBox.BackColor = System.Drawing.Color.Crimson;
            this.CatalogPathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.CatalogPathTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CatalogPathTextBox.Location = new System.Drawing.Point(141, 33);
            this.CatalogPathTextBox.Name = "CatalogPathTextBox";
            this.CatalogPathTextBox.ReadOnly = true;
            this.CatalogPathTextBox.Size = new System.Drawing.Size(231, 22);
            this.CatalogPathTextBox.TabIndex = 0;
            // 
            // OKTMOPathTextBox
            // 
            this.OKTMOPathTextBox.BackColor = System.Drawing.Color.Crimson;
            this.OKTMOPathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.OKTMOPathTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.OKTMOPathTextBox.Location = new System.Drawing.Point(141, 3);
            this.OKTMOPathTextBox.Name = "OKTMOPathTextBox";
            this.OKTMOPathTextBox.ReadOnly = true;
            this.OKTMOPathTextBox.Size = new System.Drawing.Size(231, 22);
            this.OKTMOPathTextBox.TabIndex = 1;
            // 
            // SubjectSourcePathTextBox
            // 
            this.SubjectSourcePathTextBox.BackColor = System.Drawing.Color.Crimson;
            this.SubjectSourcePathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.SubjectSourcePathTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SubjectSourcePathTextBox.Location = new System.Drawing.Point(141, 63);
            this.SubjectSourcePathTextBox.Name = "SubjectSourcePathTextBox";
            this.SubjectSourcePathTextBox.ReadOnly = true;
            this.SubjectSourcePathTextBox.Size = new System.Drawing.Size(231, 22);
            this.SubjectSourcePathTextBox.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 30);
            this.label1.TabIndex = 3;
            this.label1.Text = "ОКТМО";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(3, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(132, 30);
            this.label2.TabIndex = 4;
            this.label2.Text = "Справочник";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(3, 60);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(132, 30);
            this.label3.TabIndex = 5;
            this.label3.Text = "Субъект-Источник";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(6, 3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(189, 18);
            this.label4.TabIndex = 6;
            this.label4.Text = "Вспомогательные файлы";
            // 
            // SelectCatalogFileButton
            // 
            this.SelectCatalogFileButton.BackColor = System.Drawing.Color.WhiteSmoke;
            this.SelectCatalogFileButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SelectCatalogFileButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelectCatalogFileButton.Location = new System.Drawing.Point(378, 33);
            this.SelectCatalogFileButton.Name = "SelectCatalogFileButton";
            this.SelectCatalogFileButton.Size = new System.Drawing.Size(32, 24);
            this.SelectCatalogFileButton.TabIndex = 7;
            this.SelectCatalogFileButton.Text = "...";
            this.SelectCatalogFileButton.UseVisualStyleBackColor = false;
            this.SelectCatalogFileButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // SelectSubjectSourceFileButton
            // 
            this.SelectSubjectSourceFileButton.BackColor = System.Drawing.Color.WhiteSmoke;
            this.SelectSubjectSourceFileButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SelectSubjectSourceFileButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelectSubjectSourceFileButton.Location = new System.Drawing.Point(378, 63);
            this.SelectSubjectSourceFileButton.Name = "SelectSubjectSourceFileButton";
            this.SelectSubjectSourceFileButton.Size = new System.Drawing.Size(32, 24);
            this.SelectSubjectSourceFileButton.TabIndex = 7;
            this.SelectSubjectSourceFileButton.Text = "...";
            this.SelectSubjectSourceFileButton.UseVisualStyleBackColor = false;
            this.SelectSubjectSourceFileButton.Click += new System.EventHandler(this.SelectSubjectSourceFileButton_Click);
            // 
            // SelectOKTMOFileButton
            // 
            this.SelectOKTMOFileButton.BackColor = System.Drawing.Color.WhiteSmoke;
            this.SelectOKTMOFileButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SelectOKTMOFileButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelectOKTMOFileButton.Location = new System.Drawing.Point(378, 3);
            this.SelectOKTMOFileButton.Name = "SelectOKTMOFileButton";
            this.SelectOKTMOFileButton.Size = new System.Drawing.Size(32, 24);
            this.SelectOKTMOFileButton.TabIndex = 7;
            this.SelectOKTMOFileButton.Text = "...";
            this.SelectOKTMOFileButton.UseVisualStyleBackColor = false;
            this.SelectOKTMOFileButton.Click += new System.EventHandler(this.SelectOKTMOFileButton_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Cursor = System.Windows.Forms.Cursors.Default;
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(570, 208);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.tabPage1.Controls.Add(this.WarningLabel);
            this.tabPage1.Controls.Add(this.progressBar);
            this.tabPage1.Controls.Add(this.OpenButton);
            this.tabPage1.Controls.Add(this.workbookPathTextBox);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.StartButton);
            this.tabPage1.Cursor = System.Windows.Forms.Cursors.Default;
            this.tabPage1.Location = new System.Drawing.Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(562, 180);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Основная вкладка";
            // 
            // WarningLabel
            // 
            this.WarningLabel.BackColor = System.Drawing.Color.Yellow;
            this.WarningLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.WarningLabel.Font = new System.Drawing.Font("MS Reference Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.WarningLabel.ForeColor = System.Drawing.Color.Crimson;
            this.WarningLabel.Location = new System.Drawing.Point(3, 100);
            this.WarningLabel.Name = "WarningLabel";
            this.WarningLabel.Size = new System.Drawing.Size(437, 51);
            this.WarningLabel.TabIndex = 7;
            this.WarningLabel.Text = "Во время работы программы не открывайте никакие Excel Файлы. \r\nЭто Приведёт к пол" +
    "омке.";
            this.WarningLabel.Visible = false;
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar.ForeColor = System.Drawing.Color.LimeGreen;
            this.progressBar.Location = new System.Drawing.Point(3, 154);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(556, 23);
            this.progressBar.Step = 1;
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 6;
            // 
            // OpenButton
            // 
            this.OpenButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.OpenButton.BackColor = System.Drawing.Color.SlateGray;
            this.OpenButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OpenButton.Font = new System.Drawing.Font("MS Reference Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.OpenButton.ForeColor = System.Drawing.Color.Ivory;
            this.OpenButton.Location = new System.Drawing.Point(453, 68);
            this.OpenButton.Name = "OpenButton";
            this.OpenButton.Size = new System.Drawing.Size(106, 31);
            this.OpenButton.TabIndex = 4;
            this.OpenButton.Text = "Выбрать";
            this.OpenButton.UseVisualStyleBackColor = false;
            this.OpenButton.Click += new System.EventHandler(this.OpenButton_Click);
            // 
            // workbookPathTextBox
            // 
            this.workbookPathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.workbookPathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.workbookPathTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.workbookPathTextBox.Location = new System.Drawing.Point(3, 40);
            this.workbookPathTextBox.Name = "workbookPathTextBox";
            this.workbookPathTextBox.ReadOnly = true;
            this.workbookPathTextBox.Size = new System.Drawing.Size(556, 22);
            this.workbookPathTextBox.TabIndex = 3;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(3, 21);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(142, 16);
            this.label5.TabIndex = 2;
            this.label5.Text = "Файл для обработки";
            // 
            // StartButton
            // 
            this.StartButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.StartButton.BackColor = System.Drawing.Color.SlateGray;
            this.StartButton.Enabled = false;
            this.StartButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.StartButton.Font = new System.Drawing.Font("MS Reference Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StartButton.ForeColor = System.Drawing.Color.Ivory;
            this.StartButton.Location = new System.Drawing.Point(453, 120);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(106, 31);
            this.StartButton.TabIndex = 0;
            this.StartButton.Text = "Запустить";
            this.StartButton.UseVisualStyleBackColor = false;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.tabPage2.Controls.Add(this.tableLayoutPanel1);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Location = new System.Drawing.Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(562, 180);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Настройки";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 36.74541F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 63.25459F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 148F));
            this.tableLayoutPanel1.Controls.Add(this.label8, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.CatalogPathTextBox, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.SelectSubjectSourceFileButton, 2, 2);
            this.tableLayoutPanel1.Controls.Add(this.SubjectSourcePathTextBox, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.SelectCatalogFileButton, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.OKTMOWorksheetCBox, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.SelectOKTMOFileButton, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.OKTMOPathTextBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.CatalogWorksheetCBox, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.SubjectSourceWorksheetCBox, 3, 2);
            this.tableLayoutPanel1.Controls.Add(this.SelectVGTFileButton, 2, 3);
            this.tableLayoutPanel1.Controls.Add(this.VGTPathTextBox, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.VGTWorksheetCBox, 3, 3);
            this.tableLayoutPanel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 52);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(562, 120);
            this.tableLayoutPanel1.TabIndex = 9;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label8.Location = new System.Drawing.Point(3, 90);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(132, 30);
            this.label8.TabIndex = 9;
            this.label8.Text = "Справочник ВГТ";
            // 
            // OKTMOWorksheetCBox
            // 
            this.OKTMOWorksheetCBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.OKTMOWorksheetCBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.OKTMOWorksheetCBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OKTMOWorksheetCBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.OKTMOWorksheetCBox.FormattingEnabled = true;
            this.OKTMOWorksheetCBox.Location = new System.Drawing.Point(416, 3);
            this.OKTMOWorksheetCBox.Name = "OKTMOWorksheetCBox";
            this.OKTMOWorksheetCBox.Size = new System.Drawing.Size(143, 24);
            this.OKTMOWorksheetCBox.TabIndex = 8;
            // 
            // CatalogWorksheetCBox
            // 
            this.CatalogWorksheetCBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CatalogWorksheetCBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CatalogWorksheetCBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.CatalogWorksheetCBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CatalogWorksheetCBox.FormattingEnabled = true;
            this.CatalogWorksheetCBox.Location = new System.Drawing.Point(416, 33);
            this.CatalogWorksheetCBox.Name = "CatalogWorksheetCBox";
            this.CatalogWorksheetCBox.Size = new System.Drawing.Size(143, 24);
            this.CatalogWorksheetCBox.TabIndex = 8;
            // 
            // SubjectSourceWorksheetCBox
            // 
            this.SubjectSourceWorksheetCBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SubjectSourceWorksheetCBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.SubjectSourceWorksheetCBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SubjectSourceWorksheetCBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.SubjectSourceWorksheetCBox.FormattingEnabled = true;
            this.SubjectSourceWorksheetCBox.Location = new System.Drawing.Point(416, 63);
            this.SubjectSourceWorksheetCBox.Name = "SubjectSourceWorksheetCBox";
            this.SubjectSourceWorksheetCBox.Size = new System.Drawing.Size(143, 24);
            this.SubjectSourceWorksheetCBox.TabIndex = 8;
            // 
            // SelectVGTFileButton
            // 
            this.SelectVGTFileButton.BackColor = System.Drawing.Color.WhiteSmoke;
            this.SelectVGTFileButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.SelectVGTFileButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SelectVGTFileButton.Location = new System.Drawing.Point(378, 93);
            this.SelectVGTFileButton.Name = "SelectVGTFileButton";
            this.SelectVGTFileButton.Size = new System.Drawing.Size(32, 24);
            this.SelectVGTFileButton.TabIndex = 10;
            this.SelectVGTFileButton.Text = "...";
            this.SelectVGTFileButton.UseVisualStyleBackColor = false;
            this.SelectVGTFileButton.Click += new System.EventHandler(this.SelectVGTFileButton_Click);
            // 
            // VGTPathTextBox
            // 
            this.VGTPathTextBox.BackColor = System.Drawing.Color.Crimson;
            this.VGTPathTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.VGTPathTextBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.VGTPathTextBox.Location = new System.Drawing.Point(141, 93);
            this.VGTPathTextBox.Name = "VGTPathTextBox";
            this.VGTPathTextBox.ReadOnly = true;
            this.VGTPathTextBox.Size = new System.Drawing.Size(231, 22);
            this.VGTPathTextBox.TabIndex = 11;
            // 
            // VGTWorksheetCBox
            // 
            this.VGTWorksheetCBox.Dock = System.Windows.Forms.DockStyle.Top;
            this.VGTWorksheetCBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.VGTWorksheetCBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.VGTWorksheetCBox.FormattingEnabled = true;
            this.VGTWorksheetCBox.Location = new System.Drawing.Point(416, 93);
            this.VGTWorksheetCBox.Name = "VGTWorksheetCBox";
            this.VGTWorksheetCBox.Size = new System.Drawing.Size(143, 24);
            this.VGTWorksheetCBox.TabIndex = 12;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.Location = new System.Drawing.Point(397, 33);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(98, 16);
            this.label7.TabIndex = 11;
            this.label7.Text = "Рабочий лист";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(121, 33);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(96, 16);
            this.label6.TabIndex = 10;
            this.label6.Text = "Путь к файлу";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(570, 208);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(586, 247);
            this.MinimumSize = new System.Drawing.Size(586, 247);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Рабочая область";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button SelectCatalogFileButton;
        private System.Windows.Forms.Button SelectSubjectSourceFileButton;
        private System.Windows.Forms.Button SelectOKTMOFileButton;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.TextBox workbookPathTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button OpenButton;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox OKTMOWorksheetCBox;
        private System.Windows.Forms.ComboBox CatalogWorksheetCBox;
        private System.Windows.Forms.ComboBox SubjectSourceWorksheetCBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button SelectVGTFileButton;
        private System.Windows.Forms.ComboBox VGTWorksheetCBox;
        public System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TabPage tabPage2;
        public System.Windows.Forms.TextBox VGTPathTextBox;
        public System.Windows.Forms.TextBox CatalogPathTextBox;
        public System.Windows.Forms.TextBox OKTMOPathTextBox;
        public System.Windows.Forms.TextBox SubjectSourcePathTextBox;
        internal System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label WarningLabel;
    }
}