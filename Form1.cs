using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace tp1_echantillonnage
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Application xlAppFinal = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Workbook xlWorkBookFinal;
        Excel.Worksheet xlWorkSheet;
        Excel.Worksheet xlWorkSheetFinal;
        static int TotalRowCount;
        static int TotalColumnCount;
        public Form1()
        {
            InitializeComponent();
        }

        private void BTN_ChoisirFichier_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChoixFichier = new OpenFileDialog();
            if (ChoixFichier.ShowDialog() == DialogResult.OK)
            {
                xlWorkBook = xlApp.Workbooks.Open(ChoixFichier.FileName);
                xlWorkSheet = xlWorkBook.ActiveSheet;

                int found = ChoixFichier.FileName.LastIndexOf(@"\");
                string fileName = ChoixFichier.FileName.Substring(found + 1);
                LB_NomDuFichierChoisi.Text = fileName;

                TotalRowCount = xlWorkSheet.UsedRange.Rows.Count;
                TotalColumnCount = xlWorkSheet.UsedRange.Columns.Count;

                //worksheet.Cells[i + 1, x].Value);
            }
        }

        private void BTN_Save_Click(object sender, EventArgs e)
        {
            SaveFiles();
        }

        private void SaveFiles()
        {
            FolderBrowserDialog ChoisirPath = new FolderBrowserDialog();
            object misValue = System.Reflection.Missing.Value;
            if (ChoisirPath.ShowDialog() == DialogResult.OK)
            {        
                for (int x = 1; x <= Convert.ToInt32(TB_NbEchantillons.Text); x++)
                {
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    ModeAleatoireSimple();
                    //ModeAleatoireSimple(xlWorkSheet);
                    //RemplirWorksheet(xlWorkSheet);

                    xlWorkBookFinal.SaveAs(ChoisirPath.SelectedPath + "\\" + TB_NomsFichiers.Text + x, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            xlWorkBookFinal.Close(true, misValue, misValue);
            xlAppFinal.Quit();
        }

        //private void ModeAleatoireSimple(Excel.Worksheet worksheet)
        private void ModeAleatoireSimple()
        {
            Random random = new Random();
            var tableauRangees = new List<int> { };
            var tableauEchantillon = new List<int> { };
            // Remplit la liste des éléments
            for (int i = 0; i < TotalRowCount; i++)
            {
                tableauRangees.Add(i);
            }
            //for (int i = 0; i < Convert.ToInt32(TB_NbEchantillons.Text); i++)
            //{
                for (int j = 0; j < Convert.ToInt32(TB_TailleEchantillons.Text); j++)
                {
                    int index = random.Next(TotalRowCount);
                    //var rangee = tableauRangees[index];
                    tableauEchantillon.Add(tableauRangees[index]);
                    tableauRangees.RemoveAt(index);
                }
                // écrire dans XL
                for (int k = 0; k < tableauEchantillon.Count; k++)
                {
                    //Excel.Application xlApp2 = new Excel.Application();
                    //Excel.Workbook xlWorkBook2;
                    //Excel.Worksheet xlEchantillon;
                    object misValue = System.Reflection.Missing.Value;
                    //xlWorkBook2 = xlApp2.Workbooks.Add(misValue);
                    xlWorkBookFinal = xlAppFinal.Workbooks.Add(misValue);
                    xlWorkSheetFinal = xlWorkBookFinal.Worksheets.Add();

                    for (int l = 0; l <= TotalColumnCount-1/*Nombre de colonnes - 1*/; l++)
                    {
                        int rangeeWS = tableauEchantillon[k];
                        xlWorkSheetFinal.Cells[k + 1, l + 1] = xlWorkSheet.Cells[rangeeWS, l + 1];
                    }
                    //xlWorkBookFinal.Close(true, misValue, misValue);
                    //xlApp.Quit();
                }
            //}
        }
    }
}
