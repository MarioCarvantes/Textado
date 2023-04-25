using System;
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

        private void btnBucar_Click(object sender, EventArgs e)
        {


            ofdSeleccionar.RestoreDirectory = true;
            ofdSeleccionar.Filter = "Office Files (*.docx)|*.docx|All Files (*.*)|*.*";

            if (ofdSeleccionar.ShowDialog() == DialogResult.OK)
            {
                txtNombre.Text = ofdSeleccionar.FileName;

                txtResultado.Clear();

                object readOnly = false;
                object visible = true;
                object save = true;
                object fileName = ofdSeleccionar.FileName;
                object newTemplate = false;
                object docType = 0;
                object missing = Type.Missing;
                Microsoft.Office.Interop.Word._Document document;
                Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };
                document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                document.ActiveWindow.Selection.WholeStory();
                document.ActiveWindow.Selection.Copy();
                IDataObject dataObject = Clipboard.GetDataObject();
                rtfData.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
                application.Quit(ref missing, ref missing, ref missing);

                // Abrir el archivo seleccionado en Word
                Process.Start("WINWORD.EXE", fileName.ToString());
            }





            //ofdSeleccionar.RestoreDirectory = true;
            //ofdSeleccionar.Filter = "Office Files (*.docx)|*.docx|All Files (*.*)|*.*";



            ////Word.Application word = (Word.Application)Marshal.GetActiveObject("Word.Application");
            ////Word.Document doc = word.ActiveDocument;


            //if (ofdSeleccionar.ShowDialog() == DialogResult.OK)
            //{
            //    txtNombre.Text = ofdSeleccionar.FileName;

            //    txtResultado.Clear();

            //    object readOnly = false;
            //    object visible = true;
            //    object save = true;
            //    object fileName = ofdSeleccionar.FileName;
            //    object newTemplate = false;
            //    object docType = 0;
            //    object missing = Type.Missing;
            //    Microsoft.Office.Interop.Word._Document document;
            //    Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };
            //    document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //    document.ActiveWindow.Selection.WholeStory();
            //    document.ActiveWindow.Selection.Copy();
            //    IDataObject dataObject = Clipboard.GetDataObject();
            //    rtfData.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
            //    application.Quit(ref missing, ref missing, ref missing);

            //    Process.Start("WINWORD.EXE");




            //}

            //combinacion de los codigos utilizados, no funciona como debe, presenta el Word document pero no lo puede textar




            //using (OpenFileDialog ofdSeleccionar = new OpenFileDialog() { ValidateNames = true, Multiselect = false, Filter = "Office Files (*.docx)|*.docx|All Files (*.*)|*.* " })
            //{
            //    ofdSeleccionar.RestoreDirectory = true;

            //    if (ofdSeleccionar.ShowDialog() == DialogResult.OK)
            //    {
            //        txtNombre.Text = ofdSeleccionar.FileName;
            //        txtResultado.Clear();

            //        object readOnly = false;
            //        object visible = true;
            //        object save = true;
            //        object fileName = ofdSeleccionar.FileName;
            //        object newTemplate = false;
            //        object docType = 0;
            //        object missing = Type.Missing;
            //        Microsoft.Office.Interop.Word._Document document;
            //        Microsoft.Office.Interop.Word._Application application = new Microsoft.Office.Interop.Word.Application() { Visible = false };
            //        document = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //        document.ActiveWindow.Selection.WholeStory();
            //        document.ActiveWindow.Selection.Copy();
            //        IDataObject dataObject = Clipboard.GetDataObject();
            //        rtfData.Rtf = dataObject.GetData(DataFormats.Rtf).ToString();
            //        application.Quit(ref missing, ref missing, ref missing);
            //    }
            //}


        }

        private void BuscarHighLight(string pFileName = "")
        {
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
                wordObject.Visible = true;

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
                txtResultado.Text = Name;

                thisDoc.SaveAs2(Name, nullobject, nullobject);

                wordObject.Quit(true, Type.Missing, Type.Missing);

                circularProgress1.IsRunning = false;
                circularProgress1.Visible = false;
                if (MessageBox.Show("La versión pública del documentos se generó y guardó con éxito, " + System.Environment.NewLine + "¿Deseas abrir el directorio donde se guardó?", "Documento público guardado", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    Process Proceso = Process.Start(FilePath);
                }

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

        private void txtNombre_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtResultado_TextChanged(object sender, EventArgs e)
        {

        }

        private void ribbonPanel1_Click(object sender, EventArgs e)
        {

        }

        private void ribbonTabItem2_Click(object sender, EventArgs e)
        {

        }

        private void buttonItem1_Click(object sender, EventArgs e)
        {

        }

        private void buttonItem1_Click_1(object sender, EventArgs e)
        {

        }

        private void ofdSeleccionar_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void ribbonBar4_ItemClick(object sender, EventArgs e)
        {

        }

        private void rtfData_TextChanged(object sender, EventArgs e)
        {
                
        }

        private void labelX1_Click(object sender, EventArgs e)
        {

        }


        //boton de buscar
        private void buttonX1_Click(object sender, EventArgs e)
        {

            //codigo de la busqueda
            string[] words = txtSearch.Text.Split(',');
            foreach (string word in words)
            {
                int startIndex = 0;
                while (startIndex < rtfData.TextLength)
                {
                    int wordStartIndex = rtfData.Find(word, startIndex, RichTextBoxFinds.None);
                    if (wordStartIndex != -1)
                    {
                        rtfData.SelectionStart = wordStartIndex;
                        rtfData.SelectionLength = word.Length;
                        rtfData.SelectionBackColor = Color.Yellow ;
                    }
                    else
                        break;
                    startIndex += wordStartIndex + word.Length;
                }
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

        }

     

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        //guardado del documento, crea un documento nuevo.(buscando la manera de que se guarde en el mismo documento)
        private void guardarToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //// Crear una instancia de Word
            //Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            //// Crear un nuevo documento de Word
            //Microsoft.Office.Interop.Word.Document doc = word.Documents.Add();

            //// Agregar el contenido del RichTextBox al documento
            //string rtfData = rtfbox.SelectedRtf; // Obtener el texto RTF del RichTextBox
            //Clipboard.SetText(rtfData, TextDataFormat.Rtf); // Copiar el texto RTF al portapapeles
            //Microsoft.Office.Interop.Word.Range range = doc.Range();
            //range.Paste(); // Pegar el texto RTF en el documento de Word

            //// Guardar el documento como un archivo de Word
            //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    doc.SaveAs(saveFileDialog1.FileName);
            //}

            //// Cerrar el documento y Word
            //doc.Close();
            //word.Quit();


            // Crear una instancia de Word
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            // Crear un nuevo documento de Word
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Add();

            // Agregar el contenido del RichTextBox al documento
            Microsoft.Office.Interop.Word.Range range = doc.Range(0, 0);
            range.Text = rtfData.Text;

            // Guardar el documento como un archivo de Word
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                doc.SaveAs(saveFileDialog1.FileName);
            }

            // Cerrar el documento y Word
            doc.Close();
            word.Quit();


            //no funciona (la idea era guardar el los cambios del richTextBox en el mismo documento de word pero no funciona el "Globlas.ThisAdding" apesar de que tiene el interop de office)-
            //Microsoft.Office.Interop.Word.Application word = Globals.ThisAddIn.Application;

            //// Obtener el documento actualmente abierto
            //Microsoft.Office.Interop.Word.Document doc = word.ActiveDocument;

            //// Guardar los cambios en el documento
            //doc.Save();

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void abrirToolStripMenuItem1_Click(object sender, EventArgs e)
        {



        }


    }






}


