using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;


namespace _3CX.Funzioni
{
    internal class Algoritmo
    {
        private String passwordChar = "ABCDEFGHIJKLMNOPQRSTUVXYZ";
        private String passwordPin = "0123456789";
        private Excel.Worksheet PaginaInterni;
        private Excel.Worksheet PaginaContatti;
        private Excel.Worksheet PaginaImpostazioni;
        private Excel.Worksheet PaginaUscitaInterni;
        private Excel.Worksheet PaginaUscitaContatti;
        private readonly object[] templateInterno =  { "", "", "", "", "", "", "", "", "", "", "", "", "", "1", "", "0", "", "0", "0", "1", "", "", "0", "1", "1", "0", "0", "", "", "", "", "", "", "0", "", "", "0", "0", "1", "", "0", "", "", "1", "0", "", "", "", "", "", "", "", "", "0", "", "", "", "", "", "1", "", "1", "", "", "", "" } ;
        private Object[,] DatiCelleInterni;
        private Object[,] DatiCelleContatti;
        private Object[,] DatiCelleImpostazioni;
        private Object[,] DatiCelleUscitaInterni;
        private Object[,] DatiCelleUscitaContatti;
        private Object[,] datiImpostazioni2;

        Random randomGen = new Random();

        public string genPassword(bool PinMode = false)
        {
            string dataOut = "";
            string passSet;
            int length;

            if (PinMode)
            {
                passSet = passwordPin;
                length = 5;
            }
            else
            {
                passSet = passwordPin + passwordChar;
                length = 11;
            }

            for (int i = 0; i < length; i++)
            {
                if(randomGen.Next(2) == 1) 
                {
                    dataOut += Char.ToLower(passSet[randomGen.Next(passSet.Length)]);
                }
                else
                {
                    dataOut += passSet[randomGen.Next(passSet.Length)];
                }
                
            }
            return dataOut;
        }
        public void genPagine(System.Windows.Forms.ComboBox listaProduttori)
        {
            PaginaUscitaContatti = genWorksheet("Uscita Contatti");
            PaginaUscitaInterni = genWorksheet("Uscita Interni");
            PaginaImpostazioni = genWorksheet("Impostazioni");
            PaginaContatti = genWorksheet("Contatti");
            PaginaInterni = genWorksheet("Interni");

            PaginaUscitaContatti.Range["A:N"].EntireColumn.NumberFormat = "@";
            PaginaUscitaInterni.Range["A:BM"].EntireColumn.NumberFormat = "@";
            PaginaImpostazioni.Range["A:J"].EntireColumn.NumberFormat = "@";
            PaginaContatti.Range["A:I"].EntireColumn.NumberFormat = "@";
            PaginaInterni.Range["A:I"].EntireColumn.NumberFormat = "@";

            DatiCelleUscitaContatti = leggiFile(global::_3CX.Properties.Resources.cellaUscitaContatti);
            DatiCelleUscitaInterni = leggiFile(global::_3CX.Properties.Resources.celleUscitaInterni);
            DatiCelleImpostazioni = leggiFile(global::_3CX.Properties.Resources.celleImpostazioni);
            DatiCelleContatti = leggiFile(global::_3CX.Properties.Resources.celleContatti);
            DatiCelleInterni = leggiFile(global::_3CX.Properties.Resources.celleInterni);


            PaginaUscitaContatti.Range["A1:N" + DatiCelleUscitaContatti.GetUpperBound(0)].Value2 = DatiCelleUscitaContatti;
            PaginaUscitaInterni.Range["A1:BM" + DatiCelleUscitaInterni.GetUpperBound(0)].Value2 = DatiCelleUscitaInterni;
            PaginaImpostazioni.Range["A1:J" + DatiCelleImpostazioni.GetUpperBound(0)].Value2 = DatiCelleImpostazioni;
            PaginaContatti.Range["A1:I" + DatiCelleContatti.GetUpperBound(0)].Value2 = DatiCelleContatti;
            PaginaInterni.Range["A1:I" + DatiCelleInterni.GetUpperBound(0)].Value2 = DatiCelleInterni;
            datiImpostazioni2 = righeImpostazioni(DatiCelleImpostazioni);
            listaProduttori.Items.Clear();
            listaProduttori.Items.AddRange(Obejct2Dto1D(datiImpostazioni2));
            PaginaImpostazioni.Range["A:C"].EntireColumn.Delete();
            
        }
        public void settaTelefono(string selectedItem)
        {
            if(selectedItem != "Selezione produttore")
            {
                string query = normalizeSelection(new string[] { "G", "I" });
                Excel.Range cella = this.PaginaInterni.Range[query];

                Object[,] dataIN = cella.Value2;
                int index = 0;

                for (int i = 0; i < datiImpostazioni2.GetLength(0); i++)
                {
                    if (selectedItem == datiImpostazioni2[i, 0].ToString())
                    {
                        break;
                    }
                    index++;
                }
                String elencoTelefoni = (datiImpostazioni2[index, 1].ToString()).Replace(",", ";");

                for (int i = 0; i < dataIN.GetLength(0); i++)
                {
                    cella.Validation.Delete();
                    cella[i+1, 1].Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, elencoTelefoni, "");
                    dataIN[i + 1, 1] = "Seleziona il telefono";
                    dataIN[i + 1, 3] = datiImpostazioni2[index, 2];
                }
                cella.Value2 = dataIN;
            }
            else
            {
                MessageBox.Show("Seleziona il produttore prima di cliccare su aggiungi telefono !!!");
            }
        }
        public void genContatti()
        {
            string query = normalizeSelection(new string[] {"A", "I"});
            string query2 = normalizeSelection(new string[] { "A", "N" });

            Object[,] dataIn = this.PaginaContatti.Range[query].Value2;
            Object[,] dataOut = this.PaginaUscitaContatti.Range[query2].Value2;
            int[] offsets = { 1, 2, 10, 4, 5, 6, 7, 8, 9 };
            Object[,] dataOut2 = combineObject(dataIn, dataOut, offsets);

            this.PaginaUscitaContatti.Range[query2].Value2 = dataOut2;
           
        }
        public void genInterni()
        {
            string query1 = normalizeSelection(new string[] { "A", "H" });
            string query2 = normalizeSelection(new string[] { "A", "BM" });
            string query3 = normalizeSelection(new string[] { "B", "G" });
            
            int[] offset = { 1, 2, 3, 4, 5, 20, 54, 41 };
            int[] offset2 = { 42, 55, 57 };

            Object[,] dataIn = this.PaginaInterni.Range[query1].Value2;
            Object[,] dataOut = this.PaginaUscitaInterni.Range[query2].Value2;

            dataOut = templateInterni(dataOut);
            dataOut = combineObject(dataIn, dataOut, offset);
            
            dataOut = getSettings(dataOut);

            this.PaginaUscitaInterni.Range[query2].Value2 = dataOut;

        }

        public void settImpostazioni()
        {
            XmlDocument doc = new XmlDocument();
            XmlAttribute remoteHost =  doc.CreateAttribute("RemoteSpmHost");
            XmlAttribute remotePort = doc.CreateAttribute("RemoteSpmPort");
            string query1 = normalizeSelection(new string[] { "B", "G" });
            Object[,] sourceB;
            
            if (this.PaginaImpostazioni.Range[query1].Value2.GetType() == typeof(string))
            {
                sourceB = new Object[,] { { "", "" }, { "", this.PaginaImpostazioni.Range[query1].Value2 } };
            }
            else
            {
                sourceB = this.PaginaImpostazioni.Range[query1].Value;
            }
            for(int i = 1; i < sourceB.GetLength(0)+1; i++)
            {
                doc.LoadXml(sourceB[i, 6].ToString());
                XmlNodeList padre = doc.GetElementsByTagName("PhoneDevice");
                if (sourceB[i,2] != null)
                {

                    remoteHost.Value = sourceB[i,2].ToString();
                    remotePort.Value = "5060";
                    padre[0].Attributes["ProvType"].Value = "3";
                    padre[0].Attributes.Append(remoteHost);
                    padre[0].Attributes.Append(remotePort);
                }
                XmlNodeList figlio = padre[0].SelectNodes("//option");
                if(figlio.Count > 0 && sourceB[i, 3].ToString() != "")
                {
                    figlio[0].Attributes["value"].Value = "true";
                    figlio[1].Attributes["value"].Value = sourceB[i, 3].ToString();
                }
                if (figlio.Count > 0 && sourceB[i, 4].ToString() != "")
                {
                    figlio[3].Attributes["value"].Value = "true";
                    figlio[4].Attributes["value"].Value = sourceB[i, 4].ToString();
                }
                sourceB[i, 6]= padre[0].OuterXml;
            }
            this.PaginaImpostazioni.Range[query1].Value2 = sourceB;
        }

        //ALGORITMO PER GENERARE PAGINE SE NON ESISTONO LEGGENDO I FILE EXCEL TEMPLATE
        private Object[,] leggiFile(String File)
        {
            string[] rows = File.Split('\n');
            int colNum = (rows[0].Split(';')).Length;

            Object[,] tmpArr = new Object[rows.Length, colNum];
            for (int i = 0; i < rows.Length; i++)
            {
                string[] tmpCol = rows[i].Split(';');
                for (int j = 0; j < tmpCol.Length; j++)
                {
                    tmpArr[i, j] = tmpCol[j];
                }
            }
            return tmpArr;
        }
        private Excel.Worksheet genWorksheet(string nomePagina)
        {
            if (ChekIfExist(nomePagina))
            {
                return GetWorksheetByName(nomePagina);
            }
            Excel.Worksheet tmpPagina;
            tmpPagina = (Excel.Worksheet)Globals.main.Application.ActiveWorkbook.Worksheets.Add();
            tmpPagina.Name = nomePagina;
            return tmpPagina;

        }
        private string[] ListSheets()
        {
            string nomi = "";
            foreach (Excel.Worksheet displayWorksheet in Globals.main.Application.ActiveWorkbook.Worksheets)
            {
                nomi += displayWorksheet.Name + ",";
            }
            return nomi.Split(',');
        }
        private bool ChekIfExist(string nomePagina)
        {
            string[] lista = ListSheets();
            foreach (string s in lista)
            {
                if (s == nomePagina)
                {
                    return true;
                }
            }
            return false;
        }
        private Excel.Worksheet GetWorksheetByName(String name)
        {
            return (Excel.Worksheet)Globals.main.Application.ActiveWorkbook.Sheets[name];
        }
        private object[,] righeImpostazioni(object[,] rawdata)
        {
            Object[,] dataOut = new Object[3,3];
            for(int i = 0; i < 3; i++)
            {
                dataOut[i,0] = rawdata[i+1, 0];
                dataOut[i,1] = rawdata[i+1, 1];
                dataOut[i,2] = rawdata[i+1, 2];

            }

            return dataOut;
        }
        private Object[] Obejct2Dto1D(Object[,] rawdata, int TargetColumn = 0)
        {
            Object[] outData = new object[rawdata.GetLength(0)];

            for(int i = 0; i < rawdata.GetLength(0); i++)
            {
                outData[i] = rawdata[i, TargetColumn];
            }

            return outData;
        }
        private string normalizeSelection(string[] collums)
        {
            Excel.Range selezionato = Globals.main.Application.ActiveWindow.RangeSelection;
            string[] rowdata = (selezionato.AddressLocal).Split(':');
            string outdata = "";
            if (rowdata.Length > 1)
            {
                outdata = collums[0] + (rowdata[0].Split('$'))[2] + ":" + collums[1] + (rowdata[1].Split('$'))[2];
            }
            else
            {
                outdata = collums[0] + (rowdata[0].Split('$'))[2] + ":" + collums[1] + (rowdata[0].Split('$'))[2];
            }
            return outdata.Replace(';', ' ');
        }

        private Object[,] combineObject(Object[,] sourceA, Object[,] sourceB, int[] offsets)
        {
            
            for(int i = 0; i < sourceA.GetLength(0); i++)
            {
                for(int j = 0; j < sourceA.GetLength(1); j++)
                {
                    sourceB[i+1, offsets[j]] = sourceA[i+1, j+1];
                }
            }
            return sourceB;
        }
        private Object[,] templateInterni(Object[,] sourceA)
        {
            for(int i = 1; i  < sourceA.GetLength(0)+1; i++)
            {
                for(int j = 1; j < sourceA.GetLength(1)+1; j++)
                {
                    sourceA[i, j] = templateInterno[j];
                    string Nome = "";
                    string Cognome = "";

                    if (sourceA[i, 3] != null)
                    {
                        Cognome = sourceA[i, 3].ToString().ToLower();
                    }
                    if (sourceA[i, 2] != null)
                    {
                        Nome = sourceA[i, 2].ToString().ToLower();
                    }
                    if (Nome == "" && Cognome == "")
                    {
                        Nome = "webmetting";
                    }
                    string Pin = genPassword(true);
                    sourceA[i, 6] = genPassword();
                    sourceA[i, 7] = genPassword();
                    sourceA[i, 49] = genPassword();
                    sourceA[i, 50] = genPassword();
                    sourceA[i, 16] = Pin;
                    sourceA[i, 8] = Nome + Cognome + Pin;
                }
            }
            return sourceA;
        }
        private Object[,] getSettings(Object[,] sourceA)
        {

            string query1 = normalizeSelection(new string[] { "I", "I" });
            Object[,] sourceB;
            if(this.PaginaInterni.Range[query1].Value2 != null) { 
                if (this.PaginaInterni.Range[query1].Value2.GetType() == typeof(string))
                {
                    sourceB = new Object[,] { { "", "" }, { "" , this.PaginaInterni.Range[query1].Value2 } };
                }
                else
                {
                    sourceB = this.PaginaInterni.Range[query1].Value;
                }
                for (int i = 1; i < sourceA.GetLength(0) + 1; i++)
                {
                    if (sourceB[i, 1] == null)
                    {
                        sourceA[i, 42] = "";
                        sourceA[i, 55] = "";
                        sourceA[i, 57] = "";
                    }
                    else
                    {
                        Object[,] tmpData = this.PaginaImpostazioni.Range[$"B{sourceB[i, 1]}:G{sourceB[i, 1]}"].Value2;
                        sourceA[i, 42] = tmpData[1, 1];
                        sourceA[i, 55] = tmpData[1, 5];
                        sourceA[i, 57] = tmpData[1, 6];
                    }

                }
            }
            return sourceA;
        }

    }
}
