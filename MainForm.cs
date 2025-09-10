using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CheckSubsectionByOS_TUSUR
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            openInputPathTextBox.Text = inputPath;
            openOutPutPathTextBox.Text = outputPath;
        }

        string inputPath = "F:\\input";
        string outputPath = "F:\\output";
        private void openInputButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = openInputPathTextBox.Text;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                if (folderBrowserDialog.SelectedPath != openOutPutPathTextBox.Text)
                {
                    openInputPathTextBox.Text = folderBrowserDialog.SelectedPath;
                }
                else
                {
                    MessageBox.Show("Пути не должны совпадать!");
                }
            }
        }

        private void openOutputButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = openOutPutPathTextBox.Text;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                if (folderBrowserDialog.SelectedPath != openInputPathTextBox.Text)
                {
                    openOutPutPathTextBox.Text = folderBrowserDialog.SelectedPath;
                }
                else
                {
                    MessageBox.Show("Пути не должны совпадать!");
                }
            }
        }

        private void startProcessButton_Click(object sender, EventArgs e)
        {
            if (!System.IO.Directory.Exists(openInputPathTextBox.Text))
            {
                MessageBox.Show("Директория сканирования не существует!");
                return;
            }

            if (!System.IO.Directory.Exists(openOutPutPathTextBox.Text))
            {
                MessageBox.Show("Директория сохранения не существует!");
                return;
            }

            var files = System.IO.Directory.GetFiles(openOutPutPathTextBox.Text, "*.docx", SearchOption.AllDirectories);

            try
            {
                if (files.Length != 0 && clearDirectoryCheckbox.Checked)
                {
                    foreach (var file in files)
                    {
                        System.IO.File.Delete(file);
                    }
                }
            }
            catch 
            {
                MessageBox.Show("Не могу очистить директорию сохранения, возможно открыт файл!");
            }


            files = System.IO.Directory.GetFiles(openInputPathTextBox.Text, "*.docx", SearchOption.AllDirectories);

            try
            {
                if (files.Length != 0 )
                {
                    foreach (var file in files)
                    {
                        System.IO.File.Copy(file, openOutPutPathTextBox.Text + "\\" + new System.IO.FileInfo(file).Name, true);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Не могу скопировать документы в директорию сохранения, возможно открыт файл!");
            }

            files = System.IO.Directory.GetFiles(openOutPutPathTextBox.Text, "*.docx", SearchOption.AllDirectories);

            if (files.Length != 0)
            {
                foreach (var file in files)
                {
                    try
                    {
                        DocumentCheckUp.checkDocument(file);
                    }
                    catch (Exception exp)
                    {
                        
                        MessageBox.Show("Проблема с обработкой документа:"+ exp.Message);

                        throw;
                    }
                }

                MessageBox.Show("Проверка завершена!");
            }
            else
            {
                MessageBox.Show("Документы формата docx не найдены!");
            }
 
        }
    }
}
