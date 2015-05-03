using System;
using System.Windows.Forms;

namespace PowerpointGenerater2
{
    public partial class Instellingenform : Form
    {
        Form1 hoofdformulier;
        public Instellingenform(Form1 formulier)
        {
            InitializeComponent();
            textBox1.Text = formulier.instellingen.Templateliederen;
            textBox2.Text = formulier.instellingen.Templatetheme;
            textBox3.Text = formulier.instellingen.Databasepad;
            textBox4.Text = formulier.instellingen.regelsperslide.ToString();
            textBox5.Text = formulier.instellingen.TemplateAbeeldingLied;
            textBox6.Text = formulier.instellingen.maskPath;
            checkBox1.Checked = formulier.instellingen.dubbelePuntKolom;
            hoofdformulier = formulier;
        }

        #region Eventhandlers
        private void button1_Click(object sender, EventArgs e)
        {
            //kies een bestand en sla het pad op
            String temp = KiesFile();
            if(!temp.Equals(""))
                textBox1.Text = temp;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //kies een bestand en sla het pad op
            String temp = KiesFile();
            if (!temp.Equals(""))
                textBox2.Text = temp;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //open een open window met bepaalde instellingen
            FolderBrowserDialog openFolderDialog1 = new FolderBrowserDialog();
            openFolderDialog1.Description = "Kies map van de Database";

            //return als er word geannuleerd
            if (openFolderDialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;

            textBox3.Text = openFolderDialog1.SelectedPath;
        }
        
        private void button7_Click(object sender, EventArgs e)
        {
            //kies een bestand en sla het pad op
            String temp = KiesFile("Masks bestand|*.xml");
            if (!temp.Equals(""))
                textBox6.Text = temp;
        }
        #endregion Eventhandlers
        #region Functions
        /// <summary>
        /// Kiesfile met een PowerPoint template bestand
        /// </summary>
        /// <returns>Het gekozen bestandspad</returns>
        private string KiesFile()
        {
            return this.KiesFile("Template bestanden|*.pptx;*.potx");
        }
        /// <summary>
        /// Uitkiezen van een file aan de hand van openfiledialog
        /// </summary>
        /// <param name="type">Het type bestand: "Template bestanden|*.pptx;*.potx"</param>
        /// <returns> return gekozen bestandspad</returns>
        private String KiesFile(string type)
        {
            //open een open window met bepaalde instellingen
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = type;
            openFileDialog1.Title = "Kies bestand";

            //return als er word geannuleerd
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return "";
            //return bestandspad
            return openFileDialog1.FileName;
        }

        #endregion Functions

        private void button6_Click(object sender, EventArgs e)
        {
            //kies een bestand en sla het pad op
            String temp = KiesFile();
            if (!temp.Equals(""))
                textBox5.Text = temp;
        }

        
    }
}
