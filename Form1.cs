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
        public Form1()
        {
            InitializeComponent();
        }

        private void BTN_ChoisirFichier_Click(object sender, EventArgs e)
        {
            OpenFileDialog ChoixFichier = new OpenFileDialog();
            if (ChoixFichier.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(ChoixFichier.FileName);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                int found = ChoixFichier.FileName.LastIndexOf(@"\");
                string fileName = ChoixFichier.FileName.Substring(found + 1);
                LB_NomDuFichierChoisi.Text = fileName;

                int rowsCount = worksheet.UsedRange.Rows.Count;
                TotalRowCount = rowsCount;

                //worksheet.Cells[i + 1, x].Value);
            }
        }

        private void SaveFiles()
        {
            FolderBrowserDialog ChoisirPath = new FolderBrowserDialog();

            if (ChoisirPath.ShowDialog() == DialogResult.OK)
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                for (int x = 0; x < Convert.ToInt32(TB_NbEchantillons.Text); x++)
                {
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    ModeAleatoireSimple(xlWorkSheet);
                    //RemplirWorksheet(xlWorkSheet);

                    xlWorkBook.SaveAs(ChoisirPath.SelectedPath + "\\" + TB_NomsFichiers + x, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                }
            }
        }

        private void BTN_Save_Click(object sender, EventArgs e)
        {
            SaveFiles();
        }

        private void ModeAleatoireSimple(Excel.Worksheet worksheet)
        {
            Random random = new Random();
            var tableauRangees = new List<int> { };
            var tableauEchantillon = new List<int> { };
            // Remplit la liste des éléments
            for (int i = 0; i < worksheet.UsedRange.Rows.Count; i++)
            {
                tableauRangees.Add(i);
            }
            for (int i = 0; i < Convert.ToInt32(TB_NbEchantillons.Text); i++)
            {
                for (int j = 0; j < Convert.ToInt32(TB_TailleEchantillons.Text); j++)
                {
                    int index = random.Next(tableauRangees.Count);
                    var rangee = tableauRangees[index];
                    tableauEchantillon.Add(rangee);
                    tableauRangees.RemoveAt(index);
                    // écrire dans XL
                    for (int k = 0; k <= // Nombre de rangees - 1; k++)
                    {

                        Excel.Application xlAppli1 = new Excel.Application();
                        Excel.Workbook xlWorkBook1;
                        Excel.Worksheet xlEchantillon;
                        object misValue = System.Reflection.Missing.Value;
                        xlWorkBook1 = xlAppli1.Workbooks.Add(misValue);
                        xlEchantillon = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(1);


                        for (int l = 0; l <= //Nombre de colonnes - 1; l++)
                        {
                            int rangeeWS = tableauEchantillon[k];
                            xlEchantillon.Cells[k + 1, l + 1] = worksheet.Cells[rangeeWS, l + 1];
                        }
                        xlWorkBook1.Close(true, misValue, misValue);
                        xlAppli1.Quit();
                    }
                }
            }
        }
    }
}
