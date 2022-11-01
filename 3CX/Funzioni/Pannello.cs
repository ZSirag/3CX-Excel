using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _3CX.Funzioni
{
    public partial class Pannello : UserControl
    {
        private Algoritmo Funzioni;
        public Pannello()
        {
            InitializeComponent();
            resize();
            Funzioni = new Algoritmo();

        }
        private void resize()
        {
            int customWidth = this.splitContainer1.Width - 20;
            this.btnTelono.Width = customWidth;
            this.btnGenPagine.Width = customWidth;
            this.btnGenInterni.Width = customWidth;
            this.btnGenContatti.Width = customWidth;
            this.btnImpostazioni.Width = customWidth;
            this.listaProduttori2.Size = new Size(customWidth, 24); ;
            this.titolo1.Width = customWidth;
            this.titolo2.Width = customWidth;
        }

        private void Pannello_Resize(object sender, EventArgs e)
        {
            resize();
        }

        private void btnGenPagine_Click(object sender, EventArgs e)
        {
            Funzioni.genPagine(this.listaProduttori2);
        }

        private void listaProduttori_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnTelono_Click(object sender, EventArgs e)
        {
            Funzioni.settaTelefono(this.listaProduttori2.Text);
        }

        private void btnGenContatti_Click(object sender, EventArgs e)
        {
            Funzioni.genContatti();
        }

        private void btnGenInterni_Click(object sender, EventArgs e)
        {
            Funzioni.genInterni();
        }

        private void btnImpostazioni_Click(object sender, EventArgs e)
        {
            Funzioni.settImpostazioni();
        }
    }
}
