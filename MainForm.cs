using Microsoft.Office.Interop.Word;
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
using System.Xml.Serialization;

namespace CheckSubsectionByOS_TUSUR
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            if (File.Exists("params.xml"))
            {

               XmlSerializer xmlSerializer = new XmlSerializer(typeof(Params));

                // десериализуем объект
                using (FileStream fs = new FileStream("params.xml", FileMode.OpenOrCreate))
                {
                    parameters = xmlSerializer.Deserialize(fs) as Params;
                }
            }
            openInputPathTextBox.Text = parameters.InputPath;
            openOutPutPathTextBox.Text = parameters.OutputPath;
            dateStartScanPicker.Value = parameters.ScanDateFrom;
        }
        Params parameters = new Params();
         public class Params
        {
            public  string InputPath { get; set; } = "F:\\input";
            public  string OutputPath { get; set; } = "F:\\output";

            public DateTime ScanDateFrom { get; set; } = DateTime.Now;
        }

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
                return;
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
                return;
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

                       // throw;
                    }
                }

                MessageBox.Show("Проверка завершена!");
            }
            else
            {
                MessageBox.Show("Документы формата docx не найдены!");
            }
 
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Params));

            // десериализуем объект
            using (FileStream fs = new FileStream("params.xml", FileMode.Create))
            {
                parameters.InputPath = openInputPathTextBox.Text;
                parameters.OutputPath = openOutPutPathTextBox.Text;
                parameters.ScanDateFrom = dateStartScanPicker.Value;
                xmlSerializer.Serialize(fs, parameters);
            }
        }
 private void checkFullDirButton_Click(object sender, EventArgs e)
        {
            string startDir = openInputPathTextBox.Text;

            string[] files = Directory.EnumerateFiles(startDir, "*.docx", SearchOption.AllDirectories).ToArray();

            string pattern = "__";

            foreach (string currentFile in files)
            {
                if (new System.IO.FileInfo(currentFile).CreationTime >= dateStartScanPicker.Value &&
                    !new System.IO.FileInfo(currentFile).Name.StartsWith(pattern))
                {
                    string directory = new System.IO.FileInfo(currentFile).DirectoryName;
                    string fileName = new System.IO.FileInfo(currentFile).Name;
                    string outFileName = directory + "\\" + pattern + fileName;

                    try
                    {
                        System.IO.File.Copy(currentFile, outFileName, true);
                    }
                    catch
                    {
                        MessageBox.Show("Не могу скопировать документы в директорию сохранения, возможно открыт файл!");
                        return;
                    }

                    try
                    {
                        DocumentCheckUp.checkDocument(outFileName);
                    }
                    catch (Exception exp)
                    {

                        MessageBox.Show("Проблема с обработкой документа:" + exp.Message);

                        // throw;
                    }
                }
            }


            MessageBox.Show("Проверка завершена!");

        }

        private void checkOneWorkButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "документ подраздела|*.docx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string pattern = "__";
                string currentFile = openFileDialog.FileName;
                string directory = new System.IO.FileInfo(currentFile).DirectoryName;
                string fileName = new System.IO.FileInfo(currentFile).Name;
                string outFileName = directory + "\\" + pattern + fileName;

                try
                {
                    System.IO.File.Copy(currentFile, outFileName, true);
                }
                catch
                {
                    MessageBox.Show("Не могу скопировать документы в директорию сохранения, возможно открыт файл!");
                    return;
                }

                try
                {
                    DocumentCheckUp.checkDocument(outFileName);
                    filePathTextBox.Text = outFileName;
                    MessageBox.Show("Проверка завершена!");                    
                }
                catch (Exception exp)
                {

                    MessageBox.Show("Проблема с обработкой документа:" + exp.Message);

                    // throw;
                }


                // TODO маркированный список с тире - ругается
                // на работе котляровой анны проверить размер абзацного отступа - там некорретно высчитывает см в его попугаи
            }
        }
    }
}
