namespace Converter
{
    partial class MainForm
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
            this.workbooksTypeCBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SelectedBooksListBox = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.StartButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // workbooksTypeCBox
            // 
            this.workbooksTypeCBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.workbooksTypeCBox.FormattingEnabled = true;
            this.workbooksTypeCBox.Location = new System.Drawing.Point(15, 31);
            this.workbooksTypeCBox.Name = "workbooksTypeCBox";
            this.workbooksTypeCBox.Size = new System.Drawing.Size(274, 23);
            this.workbooksTypeCBox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(180, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Тип обрабатываемых книг";
            // 
            // SelectedBooksListBox
            // 
            this.SelectedBooksListBox.FormattingEnabled = true;
            this.SelectedBooksListBox.Location = new System.Drawing.Point(15, 90);
            this.SelectedBooksListBox.Name = "SelectedBooksListBox";
            this.SelectedBooksListBox.Size = new System.Drawing.Size(274, 121);
            this.SelectedBooksListBox.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(15, 217);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(127, 27);
            this.button1.TabIndex = 3;
            this.button1.Text = "Добавить книги";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // StartButton
            // 
            this.StartButton.Location = new System.Drawing.Point(214, 217);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(75, 27);
            this.StartButton.TabIndex = 4;
            this.StartButton.Text = "Запуск";
            this.StartButton.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(12, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(123, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "Выбранные книги";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(307, 267);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.StartButton);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.SelectedBooksListBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.workbooksTypeCBox);
            this.Name = "MainForm";
            this.Text = "Настройки";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox workbooksTypeCBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox SelectedBooksListBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label label2;
    }
}