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
            this.SuspendLayout();
            // 
            // openInputButton
            // 
            this.openInputButton.Location = new System.Drawing.Point(604, 29);
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
            this.openInputPathTextBox.Location = new System.Drawing.Point(26, 31);
            this.openInputPathTextBox.Name = "openInputPathTextBox";
            this.openInputPathTextBox.Size = new System.Drawing.Size(572, 20);
            this.openInputPathTextBox.TabIndex = 1;
            this.openInputPathTextBox.Text = "F:\\input";
            // 
            // openOutPutPathTextBox
            // 
            this.openOutPutPathTextBox.Enabled = false;
            this.openOutPutPathTextBox.Location = new System.Drawing.Point(26, 71);
            this.openOutPutPathTextBox.Name = "openOutPutPathTextBox";
            this.openOutPutPathTextBox.Size = new System.Drawing.Size(572, 20);
            this.openOutPutPathTextBox.TabIndex = 2;
            this.openOutPutPathTextBox.Text = "F:\\output";
            // 
            // openOutputButton
            // 
            this.openOutputButton.Location = new System.Drawing.Point(604, 69);
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
            this.label1.Location = new System.Drawing.Point(31, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(183, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Путь до директории сканирования";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(170, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Путь до директории сохранения";
            // 
            // startProcessButton
            // 
            this.startProcessButton.Location = new System.Drawing.Point(230, 118);
            this.startProcessButton.Name = "startProcessButton";
            this.startProcessButton.Size = new System.Drawing.Size(318, 23);
            this.startProcessButton.TabIndex = 6;
            this.startProcessButton.Text = "Запустить проверку всех документов директории";
            this.startProcessButton.UseVisualStyleBackColor = true;
            this.startProcessButton.Click += new System.EventHandler(this.startProcessButton_Click);
            // 
            // clearDirectoryCheckbox
            // 
            this.clearDirectoryCheckbox.AutoSize = true;
            this.clearDirectoryCheckbox.Checked = true;
            this.clearDirectoryCheckbox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.clearDirectoryCheckbox.Location = new System.Drawing.Point(26, 97);
            this.clearDirectoryCheckbox.Name = "clearDirectoryCheckbox";
            this.clearDirectoryCheckbox.Size = new System.Drawing.Size(283, 17);
            this.clearDirectoryCheckbox.TabIndex = 7;
            this.clearDirectoryCheckbox.Text = "Очищать директорию сохранения перед анализом";
            this.clearDirectoryCheckbox.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AcceptButton = this.startProcessButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(727, 153);
            this.Controls.Add(this.clearDirectoryCheckbox);
            this.Controls.Add(this.startProcessButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.openOutputButton);
            this.Controls.Add(this.openOutPutPathTextBox);
            this.Controls.Add(this.openInputPathTextBox);
            this.Controls.Add(this.openInputButton);
            this.Name = "MainForm";
            this.Text = "Проверка работы";
            this.ResumeLayout(false);
            this.PerformLayout();

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
    }
}

