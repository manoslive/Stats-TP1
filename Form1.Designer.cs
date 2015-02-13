namespace tp1_echantillonnage
{
    partial class Form1
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TB_NomsFichiers = new System.Windows.Forms.TextBox();
            this.TB_TailleEchantillons = new System.Windows.Forms.TextBox();
            this.RB_AleatoireSimple = new System.Windows.Forms.RadioButton();
            this.RB_Systematique = new System.Windows.Forms.RadioButton();
            this.DGV_Fichier = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GB_Methodes = new System.Windows.Forms.GroupBox();
            this.GB_Fichier = new System.Windows.Forms.GroupBox();
            this.TB_NbEchantillons = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.BTN_Save = new System.Windows.Forms.Button();
            this.BTN_ChoisirFichier = new System.Windows.Forms.Button();
            this.LB_NbRangees = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.TB_NomDuFichierChoisi = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Fichier)).BeginInit();
            this.GB_Methodes.SuspendLayout();
            this.GB_Fichier.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 29);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Nom des fichiers:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 81);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(114, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Taille des échantillons:";
            // 
            // TB_NomsFichiers
            // 
            this.TB_NomsFichiers.Enabled = false;
            this.TB_NomsFichiers.Location = new System.Drawing.Point(142, 29);
            this.TB_NomsFichiers.Margin = new System.Windows.Forms.Padding(2);
            this.TB_NomsFichiers.Name = "TB_NomsFichiers";
            this.TB_NomsFichiers.Size = new System.Drawing.Size(142, 20);
            this.TB_NomsFichiers.TabIndex = 1;
            this.TB_NomsFichiers.TextChanged += new System.EventHandler(this.TB_TextChanged);
            // 
            // TB_TailleEchantillons
            // 
            this.TB_TailleEchantillons.Enabled = false;
            this.TB_TailleEchantillons.Location = new System.Drawing.Point(142, 81);
            this.TB_TailleEchantillons.Margin = new System.Windows.Forms.Padding(2);
            this.TB_TailleEchantillons.Name = "TB_TailleEchantillons";
            this.TB_TailleEchantillons.Size = new System.Drawing.Size(142, 20);
            this.TB_TailleEchantillons.TabIndex = 3;
            this.TB_TailleEchantillons.TextChanged += new System.EventHandler(this.TB_TailleEchantillonsChanged);
            // 
            // RB_AleatoireSimple
            // 
            this.RB_AleatoireSimple.AutoSize = true;
            this.RB_AleatoireSimple.Location = new System.Drawing.Point(24, 29);
            this.RB_AleatoireSimple.Margin = new System.Windows.Forms.Padding(2);
            this.RB_AleatoireSimple.Name = "RB_AleatoireSimple";
            this.RB_AleatoireSimple.Size = new System.Drawing.Size(98, 17);
            this.RB_AleatoireSimple.TabIndex = 4;
            this.RB_AleatoireSimple.TabStop = true;
            this.RB_AleatoireSimple.Text = "Aléatoire simple";
            this.RB_AleatoireSimple.UseVisualStyleBackColor = true;
            this.RB_AleatoireSimple.CheckedChanged += new System.EventHandler(this.RB_Checked);
            // 
            // RB_Systematique
            // 
            this.RB_Systematique.AutoSize = true;
            this.RB_Systematique.Location = new System.Drawing.Point(24, 49);
            this.RB_Systematique.Margin = new System.Windows.Forms.Padding(2);
            this.RB_Systematique.Name = "RB_Systematique";
            this.RB_Systematique.Size = new System.Drawing.Size(88, 17);
            this.RB_Systematique.TabIndex = 5;
            this.RB_Systematique.TabStop = true;
            this.RB_Systematique.Text = "Systématique";
            this.RB_Systematique.UseVisualStyleBackColor = true;
            this.RB_Systematique.CheckedChanged += new System.EventHandler(this.RB_Checked);
            // 
            // DGV_Fichier
            // 
            this.DGV_Fichier.AllowUserToResizeColumns = false;
            this.DGV_Fichier.AllowUserToResizeRows = false;
            this.DGV_Fichier.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.DGV_Fichier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV_Fichier.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.DGV_Fichier.Location = new System.Drawing.Point(193, 228);
            this.DGV_Fichier.Margin = new System.Windows.Forms.Padding(2);
            this.DGV_Fichier.MultiSelect = false;
            this.DGV_Fichier.Name = "DGV_Fichier";
            this.DGV_Fichier.ReadOnly = true;
            this.DGV_Fichier.RowHeadersVisible = false;
            this.DGV_Fichier.RowTemplate.Height = 28;
            this.DGV_Fichier.Size = new System.Drawing.Size(125, 148);
            this.DGV_Fichier.TabIndex = 3;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.HeaderText = "Nom des fichiers";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // GB_Methodes
            // 
            this.GB_Methodes.Controls.Add(this.RB_Systematique);
            this.GB_Methodes.Controls.Add(this.RB_AleatoireSimple);
            this.GB_Methodes.Location = new System.Drawing.Point(11, 228);
            this.GB_Methodes.Margin = new System.Windows.Forms.Padding(2);
            this.GB_Methodes.Name = "GB_Methodes";
            this.GB_Methodes.Padding = new System.Windows.Forms.Padding(2);
            this.GB_Methodes.Size = new System.Drawing.Size(163, 85);
            this.GB_Methodes.TabIndex = 4;
            this.GB_Methodes.TabStop = false;
            this.GB_Methodes.Text = "Méthodes d\'échantillonnage";
            // 
            // GB_Fichier
            // 
            this.GB_Fichier.Controls.Add(this.TB_NbEchantillons);
            this.GB_Fichier.Controls.Add(this.TB_TailleEchantillons);
            this.GB_Fichier.Controls.Add(this.label1);
            this.GB_Fichier.Controls.Add(this.label3);
            this.GB_Fichier.Controls.Add(this.label2);
            this.GB_Fichier.Controls.Add(this.TB_NomsFichiers);
            this.GB_Fichier.Location = new System.Drawing.Point(11, 90);
            this.GB_Fichier.Margin = new System.Windows.Forms.Padding(2);
            this.GB_Fichier.Name = "GB_Fichier";
            this.GB_Fichier.Padding = new System.Windows.Forms.Padding(2);
            this.GB_Fichier.Size = new System.Drawing.Size(307, 116);
            this.GB_Fichier.TabIndex = 4;
            this.GB_Fichier.TabStop = false;
            this.GB_Fichier.Text = "Information sur l\'échantillon";
            // 
            // TB_NbEchantillons
            // 
            this.TB_NbEchantillons.Enabled = false;
            this.TB_NbEchantillons.Location = new System.Drawing.Point(142, 55);
            this.TB_NbEchantillons.Margin = new System.Windows.Forms.Padding(2);
            this.TB_NbEchantillons.Name = "TB_NbEchantillons";
            this.TB_NbEchantillons.Size = new System.Drawing.Size(142, 20);
            this.TB_NbEchantillons.TabIndex = 2;
            this.TB_NbEchantillons.TextChanged += new System.EventHandler(this.TB_NbEchantillonsChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 55);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(109, 13);
            this.label3.TabIndex = 0;
            this.label3.Text = "Nombre d\'échantillon:";
            // 
            // BTN_Save
            // 
            this.BTN_Save.Enabled = false;
            this.BTN_Save.Location = new System.Drawing.Point(11, 331);
            this.BTN_Save.Margin = new System.Windows.Forms.Padding(2);
            this.BTN_Save.Name = "BTN_Save";
            this.BTN_Save.Size = new System.Drawing.Size(163, 45);
            this.BTN_Save.TabIndex = 6;
            this.BTN_Save.Text = "Sauvegarder";
            this.BTN_Save.UseVisualStyleBackColor = true;
            this.BTN_Save.Click += new System.EventHandler(this.BTN_Save_Click);
            // 
            // BTN_ChoisirFichier
            // 
            this.BTN_ChoisirFichier.Location = new System.Drawing.Point(264, 33);
            this.BTN_ChoisirFichier.Name = "BTN_ChoisirFichier";
            this.BTN_ChoisirFichier.Size = new System.Drawing.Size(31, 21);
            this.BTN_ChoisirFichier.TabIndex = 6;
            this.BTN_ChoisirFichier.Text = "[...]";
            this.BTN_ChoisirFichier.UseVisualStyleBackColor = true;
            this.BTN_ChoisirFichier.Click += new System.EventHandler(this.BTN_ChoisirFichier_Click);
            // 
            // LB_NbRangees
            // 
            this.LB_NbRangees.AutoSize = true;
            this.LB_NbRangees.Location = new System.Drawing.Point(19, 57);
            this.LB_NbRangees.Name = "LB_NbRangees";
            this.LB_NbRangees.Size = new System.Drawing.Size(0, 13);
            this.LB_NbRangees.TabIndex = 9;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(11, 395);
            this.progressBar.MarqueeAnimationSpeed = 50;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(307, 11);
            this.progressBar.TabIndex = 10;
            // 
            // TB_NomDuFichierChoisi
            // 
            this.TB_NomDuFichierChoisi.Location = new System.Drawing.Point(19, 34);
            this.TB_NomDuFichierChoisi.Name = "TB_NomDuFichierChoisi";
            this.TB_NomDuFichierChoisi.Size = new System.Drawing.Size(239, 20);
            this.TB_NomDuFichierChoisi.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 15);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Choisir un sondage";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(332, 418);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.TB_NomDuFichierChoisi);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.LB_NbRangees);
            this.Controls.Add(this.BTN_ChoisirFichier);
            this.Controls.Add(this.BTN_Save);
            this.Controls.Add(this.GB_Fichier);
            this.Controls.Add(this.GB_Methodes);
            this.Controls.Add(this.DGV_Fichier);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.DGV_Fichier)).EndInit();
            this.GB_Methodes.ResumeLayout(false);
            this.GB_Methodes.PerformLayout();
            this.GB_Fichier.ResumeLayout(false);
            this.GB_Fichier.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TB_NomsFichiers;
        private System.Windows.Forms.TextBox TB_TailleEchantillons;
        private System.Windows.Forms.RadioButton RB_AleatoireSimple;
        private System.Windows.Forms.RadioButton RB_Systematique;
        private System.Windows.Forms.DataGridView DGV_Fichier;
        private System.Windows.Forms.GroupBox GB_Methodes;
        private System.Windows.Forms.GroupBox GB_Fichier;
        private System.Windows.Forms.TextBox TB_NbEchantillons;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BTN_Save;
        private System.Windows.Forms.Button BTN_ChoisirFichier;
        private System.Windows.Forms.Label LB_NbRangees;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TextBox TB_NomDuFichierChoisi;
        private System.Windows.Forms.Label label4;
    }
}

