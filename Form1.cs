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
            if(ChoixFichier.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(ChoixFichier.FileName);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                int found = ChoixFichier.FileName.LastIndexOf(@"\");
                string fileName = ChoixFichier.FileName.Substring(found + 1);
                LB_NomDuFichierChoisi.Text = fileName;

                int rowsCount = worksheet.UsedRange.Rows.Count;

                //worksheet.Cells[i + 1, x].Value);

            }
        }

        private void BTN_Generer_Click(object sender, EventArgs e)
        {
            DGV_Fichier.Rows.Clear();
            if(RB_AleatoireSimple.Checked == true)
            {
                CreerFichiers();
            }
            else if(RB_Systematique.Checked == true)
            {
                CreerFichiers();
            }
            BTN_Save.Enabled = true;
        }

        private void CreerFichiers()
        {
            string nomsFichiers = TB_NomsFichiers.Text;
            int nbEchantillons = Convert.ToInt32( TB_NbEchantillons.Text);
            int tailleEchantillons = Convert.ToInt32( TB_TailleEchantillons.Text);

            for(int i = 0; i<nbEchantillons; i++)
            {
                DGV_Fichier.Rows.Add(nomsFichiers + " " + (i+1));
            }

            // Resize les cells du DGV
            DGV_Fichier.AutoResizeColumnHeadersHeight();
            DGV_Fichier.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders);
        }

        private void BTN_Save_Click(object sender, EventArgs e)
        {
            SaveFiles();
        }

        private void SaveFiles()
        {
            DGV_Echantillon.Rows[0].Cells[0].Value = "Allo";

            FolderBrowserDialog ChoisirPath = new FolderBrowserDialog();

            if (ChoisirPath.ShowDialog() == DialogResult.OK)
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                for (int x = 0; x < Convert.ToInt32(TB_NbEchantillons.Text); x++ )
                {
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    for (int i = 0; i <= DGV_Echantillon.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= DGV_Echantillon.ColumnCount - 1; j++)
                        {
                            DataGridViewCell cell = DGV_Echantillon[j, i];
                            xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                        }
                    }
                    xlWorkBook.SaveAs(ChoisirPath.SelectedPath + "\\" + DGV_Fichier.Rows[x].Cells[0].Value, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                }
            }
        }
    }
}
