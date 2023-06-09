﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using Application = System.Windows.Forms.Application;
using Document = Microsoft.Office.Interop.Word.Document;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using AxEDOfficeLib;

//using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;

namespace Textado
{
    public partial class frmTextado : Form
    {

        string Caracter;


        public frmTextado()
        {
            InitializeComponent();
        }
        // search and show of the documment
        private void btnBucar_Click(object sender, EventArgs e)
        {


            ofdSeleccionar.Filter = "Office Files (*.docx)|*.docx|All Files (*.*)|*.* ";
            ofdSeleccionar.RestoreDirectory = true;
            ofdSeleccionar.Multiselect = false;
           
            if (ofdSeleccionar.ShowDialog() == DialogResult.OK)
            {
                txtNombre.Text = ofdSeleccionar.FileName;
                txtResultado.Clear();
                // opening of the documment in Edraw
                axEDOffice1.Open(txtNombre.Text);

            }



        }




        //Search of the yellow underlined
        private void BuscarHighLight(string pFileName = "")
        {
            //temporadly clse of the document
            axEDOffice1.CloseDoc(txtNombre.Text);


            Word.Application wordObject = new Word.Application();
            Word.Document thisDoc;
            try
            {
                object nullobject = System.Reflection.Missing.Value;
                object file = pFileName;
                thisDoc = wordObject.Documents.Open(ref file, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject);
                string tempStrText = "", tempStr = "";
                string FilePath = "", Name = "";
                int thisStart = 0, thisEnd = 0;
                thisDoc.Activate();
                wordObject.Visible = false;

                object missing = Type.Missing;

                Word.Range objRange = thisDoc.Content;

                objRange.Find.Highlight = 1;
                objRange.Find.Forward = true;

                //objRange.Find.Text.Contains("");

                bool found = objRange.Find.Execute("", missing, missing, missing, missing, missing, true,
                    missing, missing, missing, missing, missing, missing, missing, missing);
                if (found)
                {
                    circularProgress1.Visible = true;
                    do
                    {
                        circularProgress1.IsRunning = true;
                        string cellWordText = objRange.Text.Trim();

                        if (objRange.HighlightColorIndex.ToString() == "wdYellow")
                        {
                            objRange.Select();
                            objRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                            String foo = cellWordText;
                            String[] bar = foo.Split(' ');
                            thisStart = objRange.Start;
                            thisEnd = objRange.End;
                            foreach (String foobar in bar)
                            {
                                int total = foobar.Length;
                                for (int a = 0; a < total; a++)
                                {
                                    tempStrText = tempStrText + Caracter;

                                }
                                thisStart = thisStart;
                                thisEnd = thisStart + total;
                                ReplaceText(thisDoc, cellWordText, thisStart, thisEnd, tempStrText);
                                thisStart = thisEnd + 1;
                                tempStrText = "";
                            }


                            tempStr = "";
                            tempStrText = "";
                        }

                        int intPosition = objRange.End;
                        objRange.Start = intPosition;
                    } while (objRange.Find.Execute("", missing, missing, missing, missing, missing, true,
                        missing, missing, missing, missing, missing, missing, missing, missing));
                }
                buscar_Texto(thisDoc);
                FilePath = System.IO.Path.GetDirectoryName(ofdSeleccionar.FileName);
                Name = System.IO.Path.GetFileNameWithoutExtension(ofdSeleccionar.FileName);
               // if (System.IO.File.Exists(Name))
               //     System.IO.File.Delete(Name);

               



                int cont = 1;
                string nombredic = "_Público", Extencion = ".docx", tmpnombredic = "";
                tmpnombredic = nombredic;
                while (System.IO.File.Exists(FilePath + @"\" + Name + tmpnombredic + Extencion))
                {

                    tmpnombredic = nombredic + "(" + cont.ToString() + ")";
                    cont++;
                }

              

                nombredic = tmpnombredic;
                Name = FilePath + @"\" + Name + nombredic + Extencion;

                thisDoc.SaveAs2(Name, nullobject, nullobject);
                //pFileName.Close();
                thisDoc.Close();
                txtResultado.Text = Name;

                wordObject.Quit(true, Type.Missing, Type.Missing);
               
                axEDOffice1.Open(Name);

                circularProgress1.IsRunning = false;
                circularProgress1.Visible = false;
                //if (MessageBox.Show("La versión pública del documentos se generó y guardó con éxito, " + System.Environment.NewLine + "¿Deseas abrir el directorio donde se guardó?", "Documento público guardado", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                //{
                //    Process Proceso = Process.Start(FilePath);
                //}

            }
            catch (Exception e)
            {

            }


        }



        private void ReplaceText(Word.Document pthisDoc, string pFind = "", int pStart = 0, int pEnd = 0, string pTexted = "")
        {

            object start = pStart;
            object end = pEnd;
            Word.Range rng = pthisDoc.Range(start, pEnd);
            rng.Select();
            rng.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            rng.Find.ClearFormatting();
            rng.Text = pTexted;

        }

        private void buscar_Texto(Word.Document pthisDoc)
        {
            object missing = Type.Missing;
            object start = 1;
            object end = 1;

            Word.Range objRange = pthisDoc.Content;
            objRange.Find.Text = "daño a la hacienda";
            objRange.Find.Forward = true;


            bool found = objRange.Find.Execute("daño a la hacienda", missing, missing, missing, missing, missing, true,
                missing, missing, missing, missing, missing, missing, missing, missing);
            if (found)
            {

                objRange.Select();
                objRange.HighlightColorIndex = Word.WdColorIndex.wdDarkRed;
            }
        }

        private void BuscarHighLight_old(string pFileName = "")
        {
            /*
            try
            {
            int thisStart = 0;
            int thisEnd = 0;
            string tempStr = "", tempStrText = "", FilePath = "", Name = "";
            object SaveChanges = Word.WdSaveOptions.wdSaveChanges;
            object file = pFileName;
            object nullobject = System.Reflection.Missing.Value;
        
            thisDoc = wordObject.Documents.Open(ref file, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject);

            thisDoc.Activate();
            wordObject.Visible = true;
            
            
            Word.Range myRange;
            myRange = thisDoc.Content;
           // List<string> wordHighlights = new List<string>();
            foreach (Word.Range cellWordRange in myRange.Words)
            {
                if (cellWordRange.HighlightColorIndex.ToString() == "wdNoHighlight")
                {
                    continue;
                }
                else
                {
                    thisStart = cellWordRange.Start;
                    thisEnd = cellWordRange.End;
                    string cellWordText = cellWordRange.Text.Trim();
                    if (cellWordText.Length >= 1)   // valid word length, non-whitespace
                    {
                        int total = cellWordText.Length;
                        for (int i = 0; i < total; i++)
                        {
                            tempStrText = tempStrText + "█";
                        }
                       
                        ReplaceText(cellWordText, thisStart, thisEnd, tempStrText);
                            tempStr = "";
                            tempStrText = "";
                    }
                       
                }
            }

       
            FilePath = System.IO.Path.GetDirectoryName(ofdSeleccionar.FileName);
            Name = System.IO.Path.GetFileNameWithoutExtension(ofdSeleccionar.FileName);
            Name = FilePath + @"\" + Name + "_Textado.docx";
            txtResultado.Text = FilePath + @"\" + Name + "_Textado.docx";
            thisDoc.SaveAs2(Name, nullobject, nullobject);
            
            //thisDoc.Close(null, null, null);
            wordObject.Quit();
           
                
            }
            catch (Exception j)
            {
                MessageBox.Show(j.Message);
            }
           
        */
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Caracter = comboBox1.Text.Trim();
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmTextado_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        private void btnInicio_Click(object sender, EventArgs e)
        {
            
            string FileName = this.txtNombre.Text.Trim();
            if (!string.IsNullOrWhiteSpace(FileName))
                BuscarHighLight(FileName);
        }

        private void btnLimpiar_ItemClick(object sender, EventArgs e)
        {

        }
        //old
        //private void txtNombre_TextChanged(object sender, EventArgs e)
        //{

        //}
        //old
        //private void txtResultado_TextChanged(object sender, EventArgs e)
        //{

        //}
        //old element of toolbox, 

        //private void ribbonPanel1_Click(object sender, EventArgs e)
        //{

        //}
        //old element of toolbox, 

        //private void ribbonTabItem2_Click(object sender, EventArgs e)
        //{

        //}

        private void buttonItem1_Click(object sender, EventArgs e)
        {

        }
        //old element of toolbox, 
        //private void buttonItem1_Click_1(object sender, EventArgs e)
        //{

        //}

        private void ofdSeleccionar_FileOk(object sender, CancelEventArgs e)
        {

        }
        //old element of toolbox, 

        //private void ribbonBar4_ItemClick(object sender, EventArgs e)
        //{

        //}
        //old element of toolbox, 

        //private void rtfData_TextChanged(object sender, EventArgs e)
        //{
                
        //}
        //old element of toolbox, 

        //private void labelX1_Click(object sender, EventArgs e)
        //{

        //}


        //old element of toolbox, 

        //private void txtSearch_TextChanged(object sender, EventArgs e)
        //{

        //}




        //old element of toolbox, 

        //private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        //{
            
        //}


        //old element of toolbox, 

        //private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        //{

        //}

        //old element of toolbox, 
        //private void abrirToolStripMenuItem1_Click(object sender, EventArgs e)
        //{



        //}

        private void axEDOffice1_NotifyCtrlReady(object sender, EventArgs e)
        {

        }

      
    }






}


