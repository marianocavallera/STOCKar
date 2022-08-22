using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using STOCKar.Properties;

namespace stockaraaaaa
{
    public partial class Form1 : Form
    {

        string INC;
        string nombre;
        string activo;
        string token;
        string headset;
        string mouse;
        string cargador;
        string monitor;
        string teclado;
        int i = 0;
        string rutao;

        public Form1()
        {
            InitializeComponent();

        }

        // Change the text in a table in a word processing document.
        public void ChangeTextInCell(string filepath)
        {


            this.INC = boxINC.Text;
            this.nombre = boxNombre.Text;
            this.activo = boxActivo.Text;
            this.token = boxToken.Text;
            this.headset = boxHeadset.Text;
            this.mouse = boxMouse.Text;
            this.cargador = boxCargador.Text;
            this.monitor = boxMonitor.Text;
            this.teclado = boxTeclado.Text;
            string complemento = "";

            string[] texts = new string[] { comboBoxActivo.Text, comboBoxModelo.Text, "HP", boxActivo.Text, snActivo.Text, "" };



            if (checkToken.Checked)
            {
                complemento += " TOKEN: " + token;
            }
            if (checkHeadset.Checked)
            {
                complemento += "\n HEADSET: " + headset;
            }
            if (checkMouse.Checked)
            {
                complemento += "\n MOUSE: " + mouse;
            }
            if (checkCargador.Checked)
            {
                complemento += "\n CARGADOR: " + cargador;
            }
            if (checkMonitor.Checked)
            {
                complemento += "\n MONITOR: " + monitor;
            }
            if (checkTeclado.Checked)
            {
                complemento += "\n  TECLADO: " + teclado;
            }
            texts[5] = complemento;


            DateTime fechaActual = DateTime.Now;
            string newWord = filepath.Replace(".docx", " - " + nombre + " - " + activo + " - " + INC + ".docx");
            File.Copy(filepath, newWord);

            // Use the file name and path passed in as an argument to 
            // open an existing document.            
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(newWord, true))
            {

                // Find the first table in the document.
                Table table =
                    doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                for (i = 0; i < texts.Length; i++)
                {
                    // Find the second row in the table.
                    TableRow row = table.Elements<TableRow>().ElementAt(i);

                    // Find the second cell in the row.
                    TableCell cell = row.Elements<TableCell>().ElementAt(1);

                    // Find the first paragraph in the table cell.
                    Paragraph p = cell.Elements<Paragraph>().First();

                    // Find the first run in the paragraph.
                    Run r = p.Elements<Run>().First();

                    // Set the text for the run.
                    Text t = r.Elements<Text>().First();
                    t.Text = texts[i];
                }

            }
            //CAMBIO DE NOMBRE
            using (WordprocessingDocument doc = WordprocessingDocument.Open(newWord, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(doc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("nome");
                docText = regexText.Replace(docText, boxNombre.Text);

                using (StreamWriter sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

      
        //REMPLAZO DE TEXTO Y DECLARACION DE RUTA
        public void button1_Click(object sender, EventArgs e)
        {
           // string rutax = @"C:\Users\Marianito\Desktop\word\New folder\Termino de Asignacion.docx";

            ChangeTextInCell(rutao);
        }
        //MODELOS
        private void comboBoxActivo_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxModelo.Items.Clear();

            if (comboBoxActivo.Text == "LAPTOP")
            {
                comboBoxModelo.Items.Add("G2 645");
                comboBoxModelo.Items.Add("G3 645");
                comboBoxModelo.Items.Add("G3 725");
                comboBoxModelo.Items.Add("G3 820");
                comboBoxModelo.Items.Add("G4 645");
                comboBoxModelo.Items.Add("G6 745");
                comboBoxModelo.Items.Add("G6 830");
                comboBoxModelo.Items.Add("G7 845");
            }

            if (comboBoxActivo.Text == "DESKTOP")
            {
                comboBoxModelo.Items.Add("G2 705");
                comboBoxModelo.Items.Add("G3 705");
                comboBoxModelo.Items.Add("G5 705");
            }
            if (comboBoxActivo.Text == "DESKTOP SWING")
            {
                comboBoxModelo.Items.Add("Z400");
                comboBoxModelo.Items.Add("Z440");
                comboBoxModelo.Items.Add("Z4");
            }
            if (comboBoxActivo.Text == "LAPTOP SWING")
            {
                comboBoxModelo.Items.Add("G2 ZBOOK");
                comboBoxModelo.Items.Add("G3 ZBOOK");
                comboBoxModelo.Items.Add("G6 ZBOOK");
                comboBoxModelo.Items.Add("G7 ZBOOK");
            }
        }
        //GUARDAR RUTA
        private void btnRuta_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblRuta.Text = openFileDialog1.FileName;
                this.rutao = lblRuta.Text;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblRuta.Text = (string)STOCKar.Properties.Settings.Default["rutaGuardada"];
        }

        private void lblRuta_TextChanged(object sender, EventArgs e)
        {
            STOCKar.Properties.Settings.Default["rutaGuardada"] = lblRuta.Text;
            STOCKar.Properties.Settings.Default.Save();
        }
    }
}





