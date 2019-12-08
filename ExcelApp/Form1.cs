/*
ExcelAppSharp
VERSION: 1.0

Input: Excel File

Output: Make some processings in the Excel file

Description: This software tool allows to make processings in the Excel file.

Developer: Nicolas CHEN
*/

using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelApp
{
    public partial class ExcelApp : Form
    {
        public ExcelApp()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ExcelDialog = new OpenFileDialog();
            ExcelDialog.Filter = "Excel Files (*.xls;*.xlsx) | *.xls;*.xlsx";
            ExcelDialog.InitialDirectory = @"C:\";
            ExcelDialog.Title = "Select your team excel";
            if (ExcelDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {                
                txtFileName.Text = ExcelDialog.FileName;
                txtFileName.ReadOnly = true;               
            }
            MessageBox.Show("Excel file loaded");
            PlayMusic();
        }     

        /// <summary>
        /// Play music
        /// </summary>
        private void PlayMusic()
        {
            WMPLib.WindowsMediaPlayer wplayer = new WMPLib.WindowsMediaPlayer();
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            wplayer.URL = Path.Combine(RunningPath, "effect.mp3");
            wplayer.controls.play();
        }

        /// <summary>
        /// Read the first cell of Excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRead_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(txtFileName.Text, 1);
            MessageBox.Show(excel.ReadCell(0, 0));
            PlayMusic();
        }

        /// <summary>
        /// Write in the first cell of Excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWrite_Click(object sender, EventArgs e)
        {
            //Write in the first cell of Excel file
            Excel excel = new Excel(txtFileName.Text, 1);
            excel.WriteToCell(0, 0, txtFirstCell.Text);
            excel.Save();
            excel.Close();
            MessageBox.Show("Writing in Excel file completed");
            PlayMusic();
        }

        /// <summary>
        /// Create new Excel file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreate_Click(object sender, EventArgs e)
        {   
            //Create New File and Worksheet
            Excel excelFile = new Excel();
            excelFile.CreateNewFile();
            excelFile.CreateNewSheet();
            string filename = textBoxFilename.Text;
            string destination = textBoxDestination.Text;
            string fullPath = Path.Combine(destination, filename);
            excelFile.SaveAs(fullPath);
            excelFile.Close();
            MessageBox.Show("Creating Excel file completed");
            PlayMusic();
        }

        /// <summary>
        /// Folder Browser Dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDestination_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxDestination.Enabled = true;
                textBoxDestination.Text = folderBrowserDialog1.SelectedPath.ToString();
            }
        }

        /// <summary>
        /// Delete file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonDeleteFile_Click(object sender, EventArgs e)
        {
            //Delete the first Worksheet
            Excel excelFile = new Excel(txtFileName.Text, 1);
            excelFile.SelectWorksheet(2);
            excelFile.WriteToCell(0, 0, "This is Sheet 2");
            excelFile.DeleteWorksheet(1); //delete the sheet 1
            excelFile.Save();
            excelFile.Close();
            MessageBox.Show("Excel file Deleted");
            PlayMusic();
        }

        /// <summary>
        /// Protect Worksheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProtect_Click(object sender, EventArgs e)
        {
            //Protect Worksheet
            Excel excelFile = new Excel(txtFileName.Text, 1);
            excelFile.ProtectSheet(textBoxPassword.Text);
            excelFile.Save();
            excelFile.Close();
            MessageBox.Show("Excel file protected");
            PlayMusic();
        }

        /// <summary>
        /// Read Multiple Cells
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWriteCells_Click(object sender, EventArgs e)
        {
            //Read Multiple Cells (Read Range)
            string[,] read = new string[30, 3];
            int number = 0;

            for (int i = 0; i < 30; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    read[i, j] = number.ToString();
                    number++;
                }
            }

            Excel excelFile = new Excel(txtFileName.Text, 1);
            excelFile.WriteRange(1, 1, 30, 3, read);
            excelFile.Save();
            excelFile.Close();
            MessageBox.Show("Writing multiple cells in the Excel file completed");
            PlayMusic();
        }
    }
}
