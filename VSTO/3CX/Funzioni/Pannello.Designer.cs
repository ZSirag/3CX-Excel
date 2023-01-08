using System.Drawing;

namespace _3CX.Funzioni
{
    partial class Pannello
    {
        /// <summary> 
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione componenti

        /// <summary> 
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare 
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.listaProduttori2 = new System.Windows.Forms.ComboBox();
            this.titolo2 = new System.Windows.Forms.TextBox();
            this.titolo1 = new System.Windows.Forms.TextBox();
            this.btnImpostazioni = new System.Windows.Forms.Button();
            this.btnGenInterni = new System.Windows.Forms.Button();
            this.btnGenContatti = new System.Windows.Forms.Button();
            this.btnGenPagine = new System.Windows.Forms.Button();
            this.btnTelono = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(56)))), ((int)(((byte)(56)))), ((int)(((byte)(56)))));
            this.splitContainer1.Panel1.BackgroundImage = global::_3CX.Properties.Resources.logo_filled;
            this.splitContainer1.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.listaProduttori2);
            this.splitContainer1.Panel2.Controls.Add(this.titolo2);
            this.splitContainer1.Panel2.Controls.Add(this.titolo1);
            this.splitContainer1.Panel2.Controls.Add(this.btnImpostazioni);
            this.splitContainer1.Panel2.Controls.Add(this.btnGenInterni);
            this.splitContainer1.Panel2.Controls.Add(this.btnGenContatti);
            this.splitContainer1.Panel2.Controls.Add(this.btnGenPagine);
            this.splitContainer1.Panel2.Controls.Add(this.btnTelono);
            this.splitContainer1.Size = new System.Drawing.Size(298, 672);
            this.splitContainer1.SplitterDistance = 99;
            this.splitContainer1.TabIndex = 0;
            // 
            // listaProduttori2
            // 
            this.listaProduttori2.BackColor = System.Drawing.Color.White;
            this.listaProduttori2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listaProduttori2.ForeColor = System.Drawing.Color.Black;
            this.listaProduttori2.FormattingEnabled = true;
            this.listaProduttori2.Location = new System.Drawing.Point(0, 199);
            this.listaProduttori2.Name = "listaProduttori2";
            this.listaProduttori2.Size = new System.Drawing.Size(287, 24);
            this.listaProduttori2.TabIndex = 8;
            this.listaProduttori2.Text = "Selezione produttore";
            // 
            // titolo2
            // 
            this.titolo2.BackColor = System.Drawing.Color.White;
            this.titolo2.Enabled = false;
            this.titolo2.Font = new System.Drawing.Font("Arial", 10F);
            this.titolo2.ForeColor = System.Drawing.Color.Black;
            this.titolo2.Location = new System.Drawing.Point(0, 170);
            this.titolo2.Name = "titolo2";
            this.titolo2.Size = new System.Drawing.Size(295, 23);
            this.titolo2.TabIndex = 7;
            this.titolo2.Text = "IMPOSTAZIONI TELEFONI";
            this.titolo2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // titolo1
            // 
            this.titolo1.BackColor = System.Drawing.Color.White;
            this.titolo1.Enabled = false;
            this.titolo1.Font = new System.Drawing.Font("Arial", 10F);
            this.titolo1.ForeColor = System.Drawing.Color.Black;
            this.titolo1.Location = new System.Drawing.Point(0, 3);
            this.titolo1.Name = "titolo1";
            this.titolo1.Size = new System.Drawing.Size(295, 23);
            this.titolo1.TabIndex = 6;
            this.titolo1.Text = "IMPOSTAZIONI PAGINE";
            this.titolo1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // btnImpostazioni
            // 
            this.btnImpostazioni.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(193)))), ((int)(((byte)(220)))));
            this.btnImpostazioni.Font = new System.Drawing.Font("Arial", 10F);
            this.btnImpostazioni.ForeColor = System.Drawing.Color.White;
            this.btnImpostazioni.Location = new System.Drawing.Point(0, 275);
            this.btnImpostazioni.Name = "btnImpostazioni";
            this.btnImpostazioni.Padding = new System.Windows.Forms.Padding(5);
            this.btnImpostazioni.Size = new System.Drawing.Size(295, 40);
            this.btnImpostazioni.TabIndex = 5;
            this.btnImpostazioni.Text = "Setta impostazioni";
            this.btnImpostazioni.UseVisualStyleBackColor = false;
            this.btnImpostazioni.Click += new System.EventHandler(this.btnImpostazioni_Click);
            // 
            // btnGenInterni
            // 
            this.btnGenInterni.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(193)))), ((int)(((byte)(220)))));
            this.btnGenInterni.Font = new System.Drawing.Font("Arial", 10F);
            this.btnGenInterni.ForeColor = System.Drawing.Color.White;
            this.btnGenInterni.Location = new System.Drawing.Point(0, 124);
            this.btnGenInterni.Name = "btnGenInterni";
            this.btnGenInterni.Padding = new System.Windows.Forms.Padding(5);
            this.btnGenInterni.Size = new System.Drawing.Size(295, 40);
            this.btnGenInterni.TabIndex = 4;
            this.btnGenInterni.Text = "Crea interni";
            this.btnGenInterni.UseVisualStyleBackColor = false;
            this.btnGenInterni.Click += new System.EventHandler(this.btnGenInterni_Click);
            // 
            // btnGenContatti
            // 
            this.btnGenContatti.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(193)))), ((int)(((byte)(220)))));
            this.btnGenContatti.Font = new System.Drawing.Font("Arial", 10F);
            this.btnGenContatti.ForeColor = System.Drawing.Color.White;
            this.btnGenContatti.Location = new System.Drawing.Point(0, 78);
            this.btnGenContatti.Name = "btnGenContatti";
            this.btnGenContatti.Padding = new System.Windows.Forms.Padding(5);
            this.btnGenContatti.Size = new System.Drawing.Size(295, 40);
            this.btnGenContatti.TabIndex = 3;
            this.btnGenContatti.Text = "Crea contatti";
            this.btnGenContatti.UseVisualStyleBackColor = false;
            this.btnGenContatti.Click += new System.EventHandler(this.btnGenContatti_Click);
            // 
            // btnGenPagine
            // 
            this.btnGenPagine.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(193)))), ((int)(((byte)(220)))));
            this.btnGenPagine.Font = new System.Drawing.Font("Arial", 10F);
            this.btnGenPagine.ForeColor = System.Drawing.Color.White;
            this.btnGenPagine.Location = new System.Drawing.Point(0, 32);
            this.btnGenPagine.Name = "btnGenPagine";
            this.btnGenPagine.Padding = new System.Windows.Forms.Padding(5);
            this.btnGenPagine.Size = new System.Drawing.Size(295, 40);
            this.btnGenPagine.TabIndex = 2;
            this.btnGenPagine.Text = "Genera pagine";
            this.btnGenPagine.UseVisualStyleBackColor = false;
            this.btnGenPagine.Click += new System.EventHandler(this.btnGenPagine_Click);
            // 
            // btnTelono
            // 
            this.btnTelono.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(86)))), ((int)(((byte)(193)))), ((int)(((byte)(220)))));
            this.btnTelono.Font = new System.Drawing.Font("Arial", 10F);
            this.btnTelono.ForeColor = System.Drawing.Color.White;
            this.btnTelono.Location = new System.Drawing.Point(0, 229);
            this.btnTelono.Name = "btnTelono";
            this.btnTelono.Padding = new System.Windows.Forms.Padding(5);
            this.btnTelono.Size = new System.Drawing.Size(295, 40);
            this.btnTelono.TabIndex = 1;
            this.btnTelono.Text = "Aggiungi Telefono";
            this.btnTelono.UseVisualStyleBackColor = false;
            this.btnTelono.Click += new System.EventHandler(this.btnTelono_Click);
            // 
            // Pannello
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "Pannello";
            this.Size = new System.Drawing.Size(298, 672);
            this.Resize += new System.EventHandler(this.Pannello_Resize);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TextBox titolo2;
        private System.Windows.Forms.TextBox titolo1;
        private System.Windows.Forms.Button btnImpostazioni;
        private System.Windows.Forms.Button btnGenInterni;
        private System.Windows.Forms.Button btnGenContatti;
        private System.Windows.Forms.Button btnGenPagine;
        private System.Windows.Forms.Button btnTelono;
        private System.Windows.Forms.ComboBox listaProduttori2;
    }
}
