﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace tp1_echantillonnage
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Worksheet xlWorkSheetFinal;
        static int TotalRowCount;
        static int TotalColumnCount;
        bool executee = false;
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
                TB_NomDuFichierChoisi.Text = fileName;

                TotalRowCount = xlWorkSheet.UsedRange.Rows.Count;
                TotalColumnCount = xlWorkSheet.UsedRange.Columns.Count;

                LB_NbRangees.Text = "Taille de la population: " + (TotalRowCount - 1).ToString();

                TB_TailleEchantillons.Enabled = true;
                TB_NomsFichiers.Enabled = true;
                TB_NbEchantillons.Enabled = true;
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
            if(!executee)
                xlWorkBook = xlApp.Workbooks.Add(misValue);
            if (ChoisirPath.ShowDialog() == DialogResult.OK)
            {        
                for (int x = 1; x <= Convert.ToInt32(TB_NbEchantillons.Text); x++)
                {
                    if(VerifReady())
                    {
                        if (RB_AleatoireSimple.Checked == true)
                            ModeAleatoireSimple();
                        else if (RB_Systematique.Checked == true)
                            ModeSystematique();
                        
                        xlWorkSheetFinal.SaveAs(ChoisirPath.SelectedPath + "\\" + TB_NomsFichiers.Text + x);
                    }
                    progressBar.Value = x * 100 / Convert.ToInt32(TB_NbEchantillons.Text);
                }
            }
            executee = true;
        }
        private bool VerifReady()
        {
            bool ready = false;

            if (TB_NbEchantillons.Text != "" && TB_NomsFichiers.Text != "" && TB_TailleEchantillons.Text != "")//Si tous les champs sont remplient
            {
                RemplirDGV();
                if (RB_AleatoireSimple.Checked == true || RB_Systematique.Checked == true)//Si au moin un radio button est checked
                    ready = true;
            }
            return ready;
        }
        private void ModeSystematique()
        {
            int intervalle = TotalRowCount / Convert.ToInt32(TB_TailleEchantillons.Text);
            Random random = new Random();
            var tableauRangees = new List<int> { };
            var tableauEchantillon = new List<int> { };
            // Remplit la liste des éléments
            for (int i = 2; i <= TotalRowCount; i++) //ne peut pas pigé ligne 1, car c'est la rangée des titres
            {
                tableauRangees.Add(i);
            }
            int index = random.Next(1, intervalle);
            // Remplit une liste qui définit les éléments pigés
            for (int j = 0; j < Convert.ToInt32(TB_TailleEchantillons.Text); j++)
            {
                    tableauEchantillon.Add(tableauRangees[index]);
                    index = index + intervalle;
            }
            // Déclare un nouveau worksheet final pour y insérer les données pigés
            xlWorkSheetFinal = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            // écrire dans XL
            for (int i = 1; i <= TotalColumnCount; i++)//première rangée (titres)
            {
                xlWorkSheetFinal.Cells[1, i] = xlWorkSheet.Cells[1, i];
            }
            for (int k = 1; k <= tableauEchantillon.Count; k++)//start à la rangé 2 car  la première rangée est dédié au titres
            {
                for (int l = 1; l <= TotalColumnCount; l++)
                {
                    int rangeeWS = tableauEchantillon[k - 1];
                    xlWorkSheetFinal.Cells[k + 1, l] = xlWorkSheet.Cells[rangeeWS, l];
                }
            }
        }
        private void ModeAleatoireSimple()
        {
            Random random = new Random();
            var tableauRangees = new List<int> { };
            var tableauEchantillon = new List<int> { };
            // Remplit la liste des éléments
            for (int i = 2; i <= TotalRowCount; i++) //ne peut pas pigé ligne 1, car c'est la rangée des titres
            {
                tableauRangees.Add(i);
            }
            // Remplit une liste qui définit les randoms pigés
            for (int j = 0; j < Convert.ToInt32(TB_TailleEchantillons.Text); j++)
            {
                int index = random.Next(0,tableauRangees.Count);
                tableauEchantillon.Add(tableauRangees[index]);
                tableauRangees.RemoveAt(index);
            }
            // Déclare un nouveau worksheet final pour y insérer les données pigés
            xlWorkSheetFinal = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            // écrire dans XL
            for (int i = 1; i <= TotalColumnCount; i++ )//première rangée (titres)
            {
                xlWorkSheetFinal.Cells[1,i] = xlWorkSheet.Cells[1,i];
            }
            for (int k = 1; k <= tableauEchantillon.Count; k++)//start à la rangé 2 car  la première rangée est dédié au titres
            {
                for (int l = 1; l <= TotalColumnCount; l++)
                {
                    int rangeeWS = tableauEchantillon[k-1];
                    xlWorkSheetFinal.Cells[k + 1, l] = xlWorkSheet.Cells[rangeeWS, l];
                }
            }
        }
        // Fonction qui permet de terminer les processus d'excel encore encours
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private void TB_TextChanged(object sender, EventArgs e)
        {
            if (!Regex.IsMatch(TB_NbEchantillons.Text, @"^[0-9]+$"))
                TB_NbEchantillons.Text = "";
            else
            {
                if (VerifReady())
                    BTN_Save.Enabled = true;
                else
                    BTN_Save.Enabled = false;
            }

        }

        private void TB_NbEchantillonsChanged(object sender, EventArgs e)
        {
            if (!Regex.IsMatch(TB_NbEchantillons.Text, @"^[0-9]+$"))
                TB_NbEchantillons.Text = "";
            else
            {
                if (VerifReady())
                    BTN_Save.Enabled = true;
                else
                    BTN_Save.Enabled = false;
            }

        }
        private void TB_TailleEchantillonsChanged(object sender, EventArgs e)
        {
            if (!Regex.IsMatch(TB_TailleEchantillons.Text, @"^[0-9]+$"))
                TB_TailleEchantillons.Text = "";
            else if (Convert.ToInt32(TB_TailleEchantillons.Text) > (TotalRowCount - 1))
            {
                TB_TailleEchantillons.Text = (TotalRowCount - 1).ToString();
                if (VerifReady())
                    BTN_Save.Enabled = true;
                else
                    BTN_Save.Enabled = false;
            }

        }
        private void RB_Checked(object sender, EventArgs e)
        {
            if (VerifReady())
                BTN_Save.Enabled = true;
            else
                BTN_Save.Enabled = false;
        }

        private void TB_TailleEchantillons_TextChanged(object sender, EventArgs e)
        {

        }
        private void RemplirDGV()
        {
            DGV_Fichier.Rows.Clear();
            int nbEchantillons = Convert.ToInt32(TB_NbEchantillons.Text);
            for(int i=1;i<=nbEchantillons;i++)
            {
                DGV_Fichier.Rows.Add(TB_NomsFichiers.Text+i.ToString());
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(executee)
            {
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkSheetFinal);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
        }
    }
}
