namespace CheckSubsectionByOS_TUSUR
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.openInputButton = new System.Windows.Forms.Button();
            this.openInputPathTextBox = new System.Windows.Forms.TextBox();
            this.openOutPutPathTextBox = new System.Windows.Forms.TextBox();
            this.openOutputButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.startProcessButton = new System.Windows.Forms.Button();
            this.clearDirectoryCheckbox = new System.Windows.Forms.CheckBox();
            this.checkFullDirButton = new System.Windows.Forms.Button();
            this.filePathTextBox = new System.Windows.Forms.TextBox();
            this.checkOneWorkButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dateStartScanPicker = new System.Windows.Forms.DateTimePicker();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // openInputButton
            // 
            this.openInputButton.Location = new System.Drawing.Point(598, 39);
            this.openInputButton.Name = "openInputButton";
            this.openInputButton.Size = new System.Drawing.Size(75, 23);
            this.openInputButton.TabIndex = 0;
            this.openInputButton.Text = "Открыть";
            this.openInputButton.UseVisualStyleBackColor = true;
            this.openInputButton.Click += new System.EventHandler(this.openInputButton_Click);
            // 
            // openInputPathTextBox
            // 
            this.openInputPathTextBox.Enabled = false;
            this.openInputPathTextBox.Location = new System.Drawing.Point(20, 41);
            this.openInputPathTextBox.Name = "openInputPathTextBox";
            this.openInputPathTextBox.Size = new System.Drawing.Size(572, 20);
            this.openInputPathTextBox.TabIndex = 1;
            this.openInputPathTextBox.Text = "F:\\input";
            // 
            // openOutPutPathTextBox
            // 
            this.openOutPutPathTextBox.Enabled = false;
            this.openOutPutPathTextBox.Location = new System.Drawing.Point(20, 81);
            this.openOutPutPathTextBox.Name = "openOutPutPathTextBox";
            this.openOutPutPathTextBox.Size = new System.Drawing.Size(572, 20);
            this.openOutPutPathTextBox.TabIndex = 2;
            this.openOutPutPathTextBox.Text = "F:\\output";
            // 
            // openOutputButton
            // 
            this.openOutputButton.Location = new System.Drawing.Point(598, 79);
            this.openOutputButton.Name = "openOutputButton";
            this.openOutputButton.Size = new System.Drawing.Size(75, 23);
            this.openOutputButton.TabIndex = 3;
            this.openOutputButton.Text = "Открыть";
            this.openOutputButton.UseVisualStyleBackColor = true;
            this.openOutputButton.Click += new System.EventHandler(this.openOutputButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(183, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Путь до директории сканирования";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(170, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Путь до директории сохранения";
            // 
            // startProcessButton
            // 
            this.startProcessButton.Location = new System.Drawing.Point(20, 115);
            this.startProcessButton.Name = "startProcessButton";
            this.startProcessButton.Size = new System.Drawing.Size(318, 56);
            this.startProcessButton.TabIndex = 6;
            this.startProcessButton.Text = "Запустить проверку всех документов директории с сохранением в директорию сохранен" +
    "ия\r\n";
            this.startProcessButton.UseVisualStyleBackColor = true;
            this.startProcessButton.Click += new System.EventHandler(this.startProcessButton_Click);
            // 
            // clearDirectoryCheckbox
            // 
            this.clearDirectoryCheckbox.AutoSize = true;
            this.clearDirectoryCheckbox.Location = new System.Drawing.Point(38, 178);
            this.clearDirectoryCheckbox.Name = "clearDirectoryCheckbox";
            this.clearDirectoryCheckbox.Size = new System.Drawing.Size(283, 17);
            this.clearDirectoryCheckbox.TabIndex = 7;
            this.clearDirectoryCheckbox.Text = "Очищать директорию сохранения перед анализом";
            this.clearDirectoryCheckbox.UseVisualStyleBackColor = true;
            // 
            // checkFullDirButton
            // 
            this.checkFullDirButton.Location = new System.Drawing.Point(353, 115);
            this.checkFullDirButton.Name = "checkFullDirButton";
            this.checkFullDirButton.Size = new System.Drawing.Size(274, 56);
            this.checkFullDirButton.TabIndex = 8;
            this.checkFullDirButton.Text = "Проверка всей директории от указанной даты с сохранением под новым именем";
            this.checkFullDirButton.UseVisualStyleBackColor = true;
            this.checkFullDirButton.Click += new System.EventHandler(this.checkFullDirButton_Click);
            // 
            // filePathTextBox
            // 
            this.filePathTextBox.Enabled = false;
            this.filePathTextBox.Location = new System.Drawing.Point(8, 21);
            this.filePathTextBox.Name = "filePathTextBox";
            this.filePathTextBox.Size = new System.Drawing.Size(572, 20);
            this.filePathTextBox.TabIndex = 9;
            this.filePathTextBox.Text = "...";
            // 
            // checkOneWorkButton
            // 
            this.checkOneWorkButton.Location = new System.Drawing.Point(586, 19);
            this.checkOneWorkButton.Name = "checkOneWorkButton";
            this.checkOneWorkButton.Size = new System.Drawing.Size(75, 23);
            this.checkOneWorkButton.TabIndex = 11;
            this.checkOneWorkButton.Text = "Проверить";
            this.checkOneWorkButton.UseVisualStyleBackColor = true;
            this.checkOneWorkButton.Click += new System.EventHandler(this.checkOneWorkButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dateStartScanPicker);
            this.groupBox1.Controls.Add(this.clearDirectoryCheckbox);
            this.groupBox1.Controls.Add(this.startProcessButton);
            this.groupBox1.Controls.Add(this.checkFullDirButton);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.openOutputButton);
            this.groupBox1.Controls.Add(this.openOutPutPathTextBox);
            this.groupBox1.Controls.Add(this.openInputPathTextBox);
            this.groupBox1.Controls.Add(this.openInputButton);
            this.groupBox1.Location = new System.Drawing.Point(12, 96);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(685, 212);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Проверка группы файлов";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkOneWorkButton);
            this.groupBox2.Controls.Add(this.filePathTextBox);
            this.groupBox2.Location = new System.Drawing.Point(12, 21);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(685, 60);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Проверка одиночного документа";
            // 
            // dateStartScanPicker
            // 
            this.dateStartScanPicker.Location = new System.Drawing.Point(380, 178);
            this.dateStartScanPicker.MinDate = new System.DateTime(2025, 9, 14, 15, 51, 0, 0);
            this.dateStartScanPicker.Name = "dateStartScanPicker";
            this.dateStartScanPicker.Size = new System.Drawing.Size(200, 20);
            this.dateStartScanPicker.TabIndex = 9;
            this.dateStartScanPicker.Value = new System.DateTime(2025, 9, 14, 15, 51, 0, 0);
            // 
            // MainForm
            // 
            this.AcceptButton = this.startProcessButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(708, 321);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "MainForm";
            this.Text = "Проверка работы";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button openInputButton;
        private System.Windows.Forms.TextBox openInputPathTextBox;
        private System.Windows.Forms.TextBox openOutPutPathTextBox;
        private System.Windows.Forms.Button openOutputButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button startProcessButton;
        private System.Windows.Forms.CheckBox clearDirectoryCheckbox;
        private System.Windows.Forms.Button checkFullDirButton;
        private System.Windows.Forms.TextBox filePathTextBox;
        private System.Windows.Forms.Button checkOneWorkButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker dateStartScanPicker;
    }
}

