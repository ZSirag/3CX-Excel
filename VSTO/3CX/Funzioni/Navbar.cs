using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _3CX.Funzioni
{
    public partial class Navbar
    {
        private Pannello Control;
        Microsoft.Office.Tools.CustomTaskPane PannelloLaterale;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Control = new Pannello();
            PannelloLaterale = Globals.main.CustomTaskPanes.Add(Control, "Elenco Comandi");
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            PannelloLaterale.Visible = true;
            PannelloLaterale.Width = 300;
        }
    }
}
