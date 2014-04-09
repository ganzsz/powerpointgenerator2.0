using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;

namespace PowerpointGenerater2
{

    class parser
    {
        #region hans
        private Form1 papa;
        private List<string> liturgie;
        private List<LiturgieItem> items = new List<LiturgieItem>();
        public string errormsg = "";

        public parser(Form1 pa, string liturgie)
            : this(pa)
        {
            this.papa = pa;
            this.liturgie = new List<string>(liturgie.Split('\n'));

        }

        /// <summary>
        /// Function to parse the liturgie
        /// </summary>
        /// <returns>Returns false if one of the slides wasn't found, then use parser.errormsg (probs)</returns>
        public bool parse()
        {
            this.liturgie.Reverse();
            bool ok = true;
            for (int i = 0; i < this.liturgie.Count; i++)
            {
                if (this.liturgie[i] != "")
                {
                    LiturgieItem a = new LiturgieItem(this.liturgie[i], papa);
                    if (a.isValid)
                        this.items.Add(a);
                    ok = ok && a.isValid;
                }
            }
            this.items.Reverse();
            return ok;

        }

        /// <summary>
        /// Just a normal presentation
        /// </summary>
        /// <param name="presentatie">The presentation to parse</param>
        /// <param name="regel">Liturgie regel we're working with</param>
        /// <param name="r">counter in this.items</param>
        /// <returns>The parsed presentation</returns>
        private _Presentation normalPresentation(_Presentation presentatie, LiturgieItem regel, int r)
        {
            foreach (Slide slide in presentatie.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoTextBox)
                    {
                        switch (shape.TextFrame.TextRange.Text.ToLower())
                        {
                            case "<lezen>":
                                shape.TextFrame.TextRange.Text = regel.Titel;
                                break;
                            case "<tekst>":
                                shape.TextFrame.TextRange.Text = regel.Titel;
                                break;
                            case "<volgende>":
                                if (r < this.items.Count() - 1)
                                    shape.TextFrame.TextRange.Text = this.items[r + 1].Aansluitend;
                                else
                                    shape.TextFrame.TextRange.Text = "";
                                break;
                            case "<1e collecte:>":
                                shape.TextFrame.TextRange.Text = "1e Collecte: " + papa.textBox3.Text;
                                break;
                            case "<2e collecte:>":
                                shape.TextFrame.TextRange.Text = "2e Collecte: " + papa.textBox4.Text;
                                break;
                            case "<voorganger:>":
                                shape.TextFrame.TextRange.Text = "Voorganger: " + papa.textBox2.Text;
                                break;
                        }
                    }
                    else if (shape.Type == MsoShapeType.msoTable)
                    {
                        if (shape.Table.Rows[1].Cells[1].Shape.TextFrame.TextRange.Text.ToLower().Equals("<liturgie>"))
                        {
                            shape.Table.Rows[1].Cells[2].Shape.TextFrame.TextRange.Text = "Liturgie";
                            items.Reverse();
                            bool eerste=true;
                            foreach (LiturgieItem t in items)
                            {
                                if (t.bordje)
                                {
                                    if (eerste)
                                    {
                                        eerste = false;
                                    }
                                    else
                                    {
                                        shape.Table.Rows.Add(2);
                                    }
                                    shape.Table.Rows[2].Cells[1].Shape.TextFrame.TextRange.Text = t.bordregel[0];
                                    
                                    if (papa.instellingen.dubbelePuntKolom)
                                    {
                                        if (shape.Table.Rows[2].Cells.Count >= 4)
                                        {
                                            shape.Table.Rows[2].Cells[2].Shape.TextFrame.TextRange.Text = t.bordregel[1];
                                            shape.Table.Rows[2].Cells[3].Shape.TextFrame.TextRange.Text = t.bordregel[2];
                                            shape.Table.Rows[2].Cells[4].Shape.TextFrame.TextRange.Text = t.bordregel[3];
                                        }
                                        else
                                        {
                                            MessageBox.Show("De liturgietabel in " + regel.ToString() + " heeft geen 4 kolommen op de tweede rij");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if (shape.Table.Rows[2].Cells.Count >= 3)
                                        {
                                            shape.Table.Rows[2].Cells[2].Shape.TextFrame.TextRange.Text = t.bordregel[1] + t.bordregel[2];
                                            shape.Table.Rows[2].Cells[3].Shape.TextFrame.TextRange.Text = t.bordregel[3];
                                        }
                                        else
                                        {
                                            MessageBox.Show("De liturgietabel in " + regel.ToString() + " heeft geen 3 kolommen op de tweede rij");
                                            break;
                                        }
                                    }
                                    
                                }
                            }
                            if (!papa.textBox5.Text.Equals(""))
                            {
                                if(!papa.textBox2.TabIndex.Equals(""))
                                {
                                    shape.Table.Rows.Add(shape.Table.Rows.Count);
                                    shape.Table.Rows[shape.Table.Rows.Count-1].Cells[1].Shape.TextFrame.TextRange.Text = "L " + papa.textBox1.Text;
                                }
                                shape.Table.Rows[shape.Table.Rows.Count].Cells[1].Shape.TextFrame.TextRange.Text = "T "+papa.textBox5.Text;
                            }
                            else if (!papa.textBox1.Text.Equals(""))
                            {
                                shape.Table.Rows[shape.Table.Rows.Count].Cells[1].Shape.TextFrame.TextRange.Text = "L " + papa.textBox1.Text;
                            }
                            items.Reverse();
                        }
                    }
                }
            }
            return presentatie;
        }

        /// <summary>
        /// Puts songtext
        /// </summary>
        /// <param name="presentatie">A liedAfbeelding template presentation</param>
        /// <param name="regel">The LiturgieItem that is being worked with</param>
        /// <param name="r">The counter in this.items</param>
        /// <param name="q">The verses position</param>
        /// <param name="i">The actual verse number</param>
        /// <param name="regel">The songtext that has to be put in</param>
        /// <returns>The presentation but with the text and other elements</returns>
        public _Presentation tekstLied(_Presentation presentatie, LiturgieItem regel, int r, int q, int i, string regels)
        {
            foreach (Slide slide in presentatie.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoTextBox)
                    {
                        switch (shape.TextFrame.TextRange.Text)
                        {
                            case "<Liturgieregel>":
                                shape.TextFrame.TextRange.Text = regel.Titel;
                                if (!regel.eenvers)
                                {
                                    for (int j = q; j < regel.verzen.Count; j++)
                                    {
                                        shape.TextFrame.TextRange.Text += regel.verzen[j] + ", ";
                                    }
                                    shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.Remove(shape.TextFrame.TextRange.Text.Length - 2);
                                }
                                break;
                            case "<Volgende>":
                                if (r < this.items.Count - 1 && q == regel.verzen.Count - 1)
                                    shape.TextFrame.TextRange.Text = this.items[(r + 1)].Aansluitend;
                                else
                                    shape.TextFrame.TextRange.Text = "";
                                break;
                            case "<liedafbeelding>":
                                break;
                            case "<Inhoud>":
                                shape.TextFrame.TextRange.Text = regels;
                                break;
                        }
                    }
                }

            }
            return presentatie;
        }

        /// <summary>
        /// Puts images in presentation
        /// </summary>
        /// <param name="presentatie">A liedAfbeelding template presentation</param>
        /// <param name="regel">The LiturgieItem that is being worked with</param>
        /// <param name="r">The counter in this.items</param>
        /// <param name="q">The verses position</param>
        /// <param name="i">The actual verse number</param>
        /// <param name="v">The image that is put in the presentation</param>
        /// <returns>The presentation but with the images</returns>
        public _Presentation liedAfbeeldingPresentatie(_Presentation presentatie, LiturgieItem regel, int r, int q, int i, int v)
        {
            foreach (Slide slide in presentatie.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoTextBox)
                    {
                        switch (shape.TextFrame.TextRange.Text)
                        {
                            case "<Liturgieregel>":
                                shape.TextFrame.TextRange.Text = regel.Titel;
                                if (!regel.eenvers)
                                {
                                    for (int j = q; j < regel.verzen.Count; j++)
                                    {
                                        shape.TextFrame.TextRange.Text += regel.verzen[j] + ", ";
                                    }
                                    shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.Remove(shape.TextFrame.TextRange.Text.Length - 2);
                                }
                                break;
                            case "<Volgende>":
                                if (r < this.items.Count - 1 && q == regel.verzen.Count - 1)
                                    shape.TextFrame.TextRange.Text = this.items[(r + 1)].Aansluitend;
                                else
                                    shape.TextFrame.TextRange.Text = "";
                                break;
                            case "<liedafbeelding>":
                                if (File.Exists(regel.psalmmap + '\\' + i + @"-" + v + ".gif"))
                                {
                                    slide.Shapes.AddPicture(regel.psalmmap + '\\' + i + @"-" + v + ".gif", MsoTriState.msoFalse, MsoTriState.msoTrue, shape.TextFrame.TextRange.BoundLeft, shape.Top, shape.Width, shape.Height);
                                    shape.TextFrame.TextRange.Text = "";
                                    shape.Width = 0.0001f;
                                    shape.Height = 0.0001f;
                                    shape.Left = -1;
                                }
                                else
                                    break;
                                break;
                            case "<Inhoud>":
                                break;
                        }
                    }
                }

            }
            return presentatie;
        }

        public void createPresentation()
        {
            if (this.objApp == null)
                return;

            for (int r = 0; r < this.items.Count; r++)
            {
                LiturgieItem regel = this.items[r];
                if (regel.isLied)
                {
                    for (int q = 0; q < regel.verzen.Count; q++)
                    {
                        _Presentation presentatie;
                        int i = regel.verzen[q];
                        #region liedafbeelding
                        if (File.Exists(regel.psalmmap + '\\' + i + @"-1.gif"))
                        {
                            for (int v = 1; v < 100; v++)
                            {
                                if (File.Exists(regel.psalmmap + '\\' + i + @"-" + v + ".gif"))
                                {
                                    if (File.Exists(papa.instellingen.TemplateAbeeldingLied))
                                    {
                                        presentatie = OpenPPS(papa.instellingen.TemplateAbeeldingLied);
                                        presentatie = this.liedAfbeeldingPresentatie(presentatie, regel, r, q, i, v);
                                    }
                                    else
                                    {
                                        MessageBox.Show("het pad naar de liedafbeeldingtemplate is niet gezet");
                                        ClosePPS();
                                        return;
                                    }
                                    VoegSlideinPresentatiein(presentatie.Slides);
                                    //sluit de template weer af
                                    presentatie.Close();
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        #endregion liedafbeelding
                        #region liedtekst
                        //TODO fix eeuwig laatste vers
                        else
                        {
                            string[] versregels;
                            try
                            {
                                //open een filestream naar het gekozen bestand
                                FileStream strm = new FileStream(regel.psalmmap + '\\' + i + ".txt", FileMode.Open, FileAccess.Read);

                                //gebruik streamreader om te lezen van de filestream
                                using (StreamReader rdr = new StreamReader(strm))
                                {
                                    //return de liturgie
                                    versregels = rdr.ReadToEnd().Split('\n');
                                    string vv = "";
                                    foreach (string tv in versregels)
                                    {
                                        if (tv != "" && tv != "\r")
                                            vv += tv + "\n";
                                    }
                                    versregels = vv.Split('\n');
                                    bool urn = true;
                                    int c = 0;
                                    while (urn)
                                    {
                                        string regels = "";
                                        for (int d = 0; d < papa.instellingen.regelsperslide; d++)
                                        {
                                            if (((c * papa.instellingen.regelsperslide) + d) < versregels.Count())
                                            {
                                                if (versregels[((c * papa.instellingen.regelsperslide) + d)] != "\r"
                                                    && versregels[((c * papa.instellingen.regelsperslide) + d)] != "")
                                                {
                                                    regels += versregels[((c * papa.instellingen.regelsperslide) + d)];
                                                    if (d == (papa.instellingen.regelsperslide - 1))
                                                    {
                                                        if (regels.EndsWith("\r"))
                                                        {
                                                            regels = regels.Remove(regels.Count() - 1);
                                                        }
                                                        regels += ">>";
                                                    }
                                                }

                                            }
                                            else
                                            {
                                                urn = false;
                                                break;
                                            }
                                        }
                                        if (File.Exists(papa.instellingen.Templateliederen))
                                        {
                                            presentatie = OpenPPS(papa.instellingen.Templateliederen);
                                            presentatie = this.tekstLied(presentatie, regel, r, q, i, regels);
                                        }
                                        else
                                        {
                                            MessageBox.Show("het pad naar de liedtemplate is niet gezet");
                                            ClosePPS();
                                            return;
                                        }
                                        VoegSlideinPresentatiein(presentatie.Slides);
                                        //sluit de template weer af
                                        presentatie.Close();
                                        c++;
                                    }



                                }
                            }
                            //vang errors af en geef een melding dat er iets is fout gegaan
                            catch (Exception)
                            {
                                MessageBox.Show("Fout tijdens openen bestand \"" + regel.psalmmap + '\\' + i + ".txt" + "\"", "Bestand error",
                                           MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                            }


                        }
                        #endregion liedtekst

                    }
                }
                else
                {
                    _Presentation presentatie = OpenPPS(regel.presentatiepad);
                    presentatie = this.normalPresentation(presentatie, regel, r);
                    VoegSlideinPresentatiein(presentatie.Slides);
                    presentatie.Close();
                }
                papa.progressBar1.PerformStep();
            }
            System.Diagnostics.Debug.WriteLine("hallo klaar");
            papa.autoEvent.Set();
            objApp.WindowState = PpWindowState.ppWindowMaximized;
            return;
        }

        
        #endregion hans
        #region oud
        private Microsoft.Office.Interop.PowerPoint.Application objApp;
        private Presentations objPresSet;
        private _Presentation objPres;
        private CustomLayout layout;
        private Form1 hoofdformulier;
        private int slideteller = 1;

        private List<Uitlezen_Liturgie> Liturgie = new List<Uitlezen_Liturgie>();
        private string templateLiederen;
        private string Voorganger;
        private string Collecte1;
        private string Collecte2;
        private string Lezen;
        private string Tekst;

        public parser(Form1 hoofdformulier)
        {
            System.Diagnostics.Debug.Print(hoofdformulier.instellingen.Databasepad);
            if (File.Exists(hoofdformulier.instellingen.Templatetheme))
            {
                //Creeer een nieuwe lege presentatie volgens een bepaald thema
                objApp = new Microsoft.Office.Interop.PowerPoint.Application();
                objApp.Visible = MsoTriState.msoTrue;
                objPresSet = objApp.Presentations;
                objPres = objPresSet.Open(hoofdformulier.instellingen.Templatetheme,
                    MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
                //sla het thema op, zodat dat in iedere nieuwe slide kan worden meegenomen
                layout = objPres.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];
                //minimaliseer powerpoint
                objApp.WindowState = PpWindowState.ppWindowMinimized;

                this.hoofdformulier = hoofdformulier;
            }
            else
                MessageBox.Show("het pad naar de achtergrond powerpoint presentatie kan niet worden gevonden.\n\n stel de achtergrond opnieuw in bij de templates", "Stel template opnieuw in", MessageBoxButtons.OK);
        }

        /// <summary>
        /// Open een presentatie op het meegegeven pad
        /// </summary>
        /// <param name="path">het pad waar de powerpointpresentatie kan worden gevonden</param>
        /// <returns>de powerpoint presentatie</returns>
        public _Presentation OpenPPS(String path)
        {
            //controleer voor het openen van de presentatie op het meegegeven path of de presentatie bestaat
            if (File.Exists(path))
            {
                //open de presentatie op de meegegeven pad
                Presentation objPres1 = objApp.Presentations.Open(path,
                    MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
                //return de geopende presentatie
                return objPres1;
            }
            return null;
        }

        public void ClosePPS()
        {
            if (objApp != null && objPres != null)
            {
                objPres.Close();
                objApp.Quit();
            }
        }

        public void InputGeneratePresentation(List<Uitlezen_Liturgie> Liturgie, string templateLiederen, string Voorganger, string Collecte1, string Collecte2, string Lezen, string Tekst)
        {
            this.Liturgie = Liturgie;
            this.templateLiederen = templateLiederen;
            this.Voorganger = Voorganger;
            this.Collecte1 = Collecte1;
            this.Collecte2 = Collecte2;
            this.Lezen = Lezen;
            this.Tekst = Tekst;
        }

        /// <summary>
        /// Genereer een presentatie aan de hand van meegegeven Liturgie en Template voor de Liederen
        /// </summary>
        /// <param name="Liturgie">Liturgie die de indeling en inhoud van de gegenereerde presentatie bepaald</param>
        public void GeneratePresentation()
        {
            if (objApp == null)
                return;


            int counterliturgieregel = 0;
            //voor elke regel in de liturgie moeten sheets worden gemaakt
            foreach (Uitlezen_Liturgie regel in Liturgie)
            {
                #region Liederen
                int counterliedtekst = 0;
                //als de regel tekst bevat moet er voor ieder vers een sheet worden gemaakt
                foreach (String liedtekst in regel.inhoud)
                {
                    String templiedtekst = liedtekst;
                    //zolang er nog tekst is in te voegen in sheets
                    while (!templiedtekst.Equals(""))
                    {
                        _Presentation presentatie;
                        if (File.Exists(templateLiederen))
                        {
                            //lees de template uit
                            presentatie = OpenPPS(templateLiederen);
                        }
                        else
                        {
                            MessageBox.Show("het pad naar de liederen template powerpoint presentatie kan niet worden gevonden\n stel de achtergrond opnieuw in bij de templates", "Template niet gevonden", MessageBoxButtons.OK);
                            ClosePPS();
                            return;
                        }
                        //voor elke slide in de presentatie(in principe moet dit er 1 zijn)
                        foreach (Slide slide in presentatie.Slides)
                        {
                            //voor elk object op de slides (we zoeken naar de tekst die vervangen moet worden in de template)
                            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                            {
                                //als de shape gelijk is aan een textbox bevat het dus tekst
                                if (shape.Type == MsoShapeType.msoTextBox)
                                {
                                    //als de template de tekst bevat "Liturgieregel" moet daar de liturgieregel komen
                                    if (shape.TextFrame.TextRange.Text.Equals("<Liturgieregel>"))
                                    {
                                        shape.TextFrame.TextRange.Text = regel.mappath + " ";
                                        if (0 == counterliedtekst)
                                        {
                                            shape.TextFrame.TextRange.Text += regel.bestandsnamen[0];
                                        }

                                        //zoek de juiste bestandsnaam op
                                        for (int i = 1; i < regel.bestandsnamen.Count; i++)
                                        {
                                            if (i == counterliedtekst)
                                            {
                                                shape.TextFrame.TextRange.Text += regel.bestandsnamen[i];
                                            }
                                            else if (i > counterliedtekst)
                                            {
                                                shape.TextFrame.TextRange.Text += ", " + regel.bestandsnamen[i];
                                            }
                                        }
                                    }
                                    //als de template de tekst bevat "Inhoud" moet daar de inhoud van het vers komen
                                    if (shape.TextFrame.TextRange.Text.Equals("<Inhoud>"))
                                    {
                                        System.Windows.Forms.RichTextBox text = new System.Windows.Forms.RichTextBox();
                                        text.Text = templiedtekst;
                                        //leeg het tekstveld
                                        shape.TextFrame.TextRange.Text = "";
                                        //leeg de variabele liedtekst
                                        templiedtekst = "";
                                        //haal maximaal regelsperslide regels van het vers op en zet de rest terug in liedtekst
                                        int counter = 0;
                                        bool NewSlide = false;
                                        foreach (String line in text.Lines)
                                        {
                                            if (!line.Equals(""))
                                            {
                                                //zet in de sheet
                                                if (counter < hoofdformulier.instellingen.regelsperslide)
                                                {
                                                    //update de tekst
                                                    shape.TextFrame.TextRange.Text += line;
                                                    if ((counter + 1) < hoofdformulier.instellingen.regelsperslide)
                                                        shape.TextFrame.TextRange.Text += "\r\n";
                                                }
                                                //zet terug in liedtekst
                                                else
                                                {
                                                    NewSlide = true;
                                                    templiedtekst += line;
                                                    templiedtekst += "\r\n";
                                                }
                                                counter++;
                                            }
                                        }
                                        if (NewSlide)
                                            shape.TextFrame.TextRange.Text += " >>";
                                    }
                                    //als de template de tekst bevat "Volgende" moet daar de Liturgieregel van de volgende sheet komen
                                    if (shape.TextFrame.TextRange.Text.Equals("<Volgende>"))
                                    {
                                        //zoek de juiste bestandsnaam op, dat is dus niet de huidige maar de volgende daarom -1
                                        int bestandsnamencounter = 0;
                                        String bestandsnaam = "";
                                        foreach (String tempstring in regel.bestandsnamen)
                                        {
                                            if ((bestandsnamencounter - 1) == counterliedtekst)
                                            {
                                                bestandsnaam = tempstring;
                                            }
                                            bestandsnamencounter++;
                                        }
                                        //als er een volgende is gevonden
                                        if (!bestandsnaam.Equals(""))
                                        {
                                            //update de tekst met het bestandsnaam (bestandsnaam is in het geval van een lied de Liturgieregel)
                                            //shape.TextFrame.TextRange.Text = "Hierna: ";
                                            //shape.TextFrame.TextRange.Text += bestandsnaam;
                                            shape.TextFrame.TextRange.Text = "";
                                        }
                                        else
                                        {
                                            //update de tekst met de volgende liturgie als die er is
                                            if (Liturgie.Count > (counterliturgieregel + 1))
                                            {
                                                //update alleen als de tekst niet blanco is, omdat het lelijk is om blanco te zien staan
                                                if (!Liturgie[counterliturgieregel + 1].bestandsnamen[0].Equals("Blanco"))
                                                {
                                                    shape.TextFrame.TextRange.Text = "Aansluitend: ";
                                                    shape.TextFrame.TextRange.Text += Liturgie[counterliturgieregel + 1].mappath + " " + Liturgie[counterliturgieregel + 1].bestandsnamen[0];
                                                    //zoek de juiste bestandsnaam op
                                                    for (int i = 1; i < Liturgie[counterliturgieregel + 1].bestandsnamen.Count; i++)
                                                    {
                                                        shape.TextFrame.TextRange.Text += ", " + Liturgie[counterliturgieregel + 1].bestandsnamen[i];
                                                    }
                                                }
                                                else
                                                    shape.TextFrame.TextRange.Text = "";
                                            }
                                            else
                                                shape.TextFrame.TextRange.Text = "";
                                        }
                                    }
                                }
                            }
                        }
                        //voeg slide in in het grote geheel
                        VoegSlideinPresentatiein(presentatie.Slides);
                        //sluit de template weer af
                        presentatie.Close();
                    }
                    counterliedtekst++;
                }
                #endregion Liederen
                #region Slides
                //als de regel kant en klare sheets bevat voegen wij deze in
                if (!regel.sheetspath.Equals(""))
                {
                    //open de presentatie met de sheets erin
                    _Presentation presentatie = OpenPPS(regel.sheetspath);
                    //voor elke slide in de presentatie(in principe moet dit er 1 zijn)
                    foreach (Slide slide in presentatie.Slides)
                    {
                        //voor elk shape in de slide (we zoeken naar de tekst of andere dingen die vervangen moet worden in de geopende sheet)
                        foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                        {
                            //als de shape gelijk is aan een textbox bevat het dus tekst
                            if (shape.Type == MsoShapeType.msoTextBox)
                            {
                                //als de template de tekst bevat "Voorganger: " moet daar de Voorgangersnaam achter komen
                                if (shape.TextFrame.TextRange.Text.Equals("<Voorganger:>"))
                                {
                                    shape.TextFrame.TextRange.Text = "Voorganger: ";
                                    shape.TextFrame.TextRange.Text += Voorganger;
                                }
                                //als de template de tekst bevat "1e Collecte: " moet daar de Voorgangersnaam achter komen
                                if (shape.TextFrame.TextRange.Text.Equals("<1e Collecte:>"))
                                {
                                    shape.TextFrame.TextRange.Text = "1e Collecte: ";
                                    shape.TextFrame.TextRange.Text += Collecte1;
                                }
                                //als de template de tekst bevat "2e Collecte: " moet daar de Voorgangersnaam achter komen
                                if (shape.TextFrame.TextRange.Text.Equals("<2e Collecte:>"))
                                {
                                    shape.TextFrame.TextRange.Text = "2e Collecte: ";
                                    shape.TextFrame.TextRange.Text += Collecte2;
                                }
                                //als de template de tekst bevat "Volgende" moet daar de Liturgieregel van de volgende sheet komen
                                if (shape.TextFrame.TextRange.Text.Equals("<Volgende>"))
                                {
                                    //update de tekst met de volgende liturgie als die er is
                                    if (Liturgie.Count > (counterliturgieregel + 1))
                                    {
                                        //update alleen als de tekst niet blanco is, omdat het lelijk is om blanco te zien staan
                                        if (!Liturgie[counterliturgieregel + 1].bestandsnamen[0].Equals("Blanco"))
                                        {
                                            shape.TextFrame.TextRange.Text = "Aansluitend: ";
                                            shape.TextFrame.TextRange.Text += Liturgie[counterliturgieregel + 1].mappath + " " + Liturgie[counterliturgieregel + 1].bestandsnamen[0];
                                            //zoek de juiste bestandsnamen op
                                            for (int i = 1; i < Liturgie[counterliturgieregel + 1].bestandsnamen.Count; i++)
                                            {
                                                shape.TextFrame.TextRange.Text += ", " + Liturgie[counterliturgieregel + 1].bestandsnamen[i];
                                            }
                                        }
                                        else
                                            shape.TextFrame.TextRange.Text = "";
                                    }
                                    else
                                        shape.TextFrame.TextRange.Text = "";
                                }
                                //als de template de tekst bevat "Volgende" moet daar de te lezen schriftgedeeltes komen
                                if (shape.TextFrame.TextRange.Text.Equals("<Lezen>"))
                                {
                                    shape.TextFrame.TextRange.Text = "Schriftlezing:\n";
                                    shape.TextFrame.TextRange.Text += Lezen;
                                }
                                if (shape.TextFrame.TextRange.Text.Equals("<Tekst>"))
                                {
                                    shape.TextFrame.TextRange.Text = "Tekst:\n";
                                    shape.TextFrame.TextRange.Text += Tekst;
                                }
                            }
                            if (shape.Type == MsoShapeType.msoTable)
                            {
                                if (shape.Table.Rows[1].Cells[1].Shape.TextFrame.TextRange.Text.Equals("<Liturgie>"))
                                {
                                    int tempcounter = 0;
                                    bool legeregel = false;
                                    bool lezengehad = false;
                                    bool tekstgehad = false;
                                    List<Row> deleterows = new List<Row>();
                                    for (int index = 1; index <= shape.Table.Rows.Count; index++)
                                    {
                                        if (!shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.Text.Equals("<Liturgie>"))
                                        {
                                            Boolean liturgiegevonden = false;
                                            while ((Liturgie.Count > tempcounter) && (!liturgiegevonden))
                                            {
                                                if (Liturgie[tempcounter].Liturgiebord)
                                                {
                                                    string[] tempdelen = Liturgie[tempcounter].mappath.Split(' ');
                                                    for (int i = 0; i < tempdelen.Count(); i++)
                                                    {
                                                        if (!tempdelen[i].Equals(" "))
                                                        {
                                                            shape.Table.Rows[index].Cells[i + 1].Shape.TextFrame.TextRange.Text = tempdelen[i];
                                                        }
                                                    }
                                                    shape.Table.Rows[index].Cells[tempdelen.Count()].Shape.TextFrame.TextRange.Text = Liturgie[tempcounter].bestandsnamen[0];
                                                    for (int i = 1; i < Liturgie[tempcounter].bestandsnamen.Count; i++)
                                                    {
                                                        shape.Table.Rows[index].Cells[tempdelen.Count()].Shape.TextFrame.TextRange.Text += ", " + Liturgie[tempcounter].bestandsnamen[i];
                                                    }
                                                    liturgiegevonden = true;
                                                }
                                                tempcounter++;
                                            }
                                            if (!liturgiegevonden)
                                            {
                                                shape.Table.Rows[index].Cells[1].Merge(shape.Table.Rows[index].Cells[2]);
                                                if (shape.Table.Rows[index].Cells.Count >= 3)
                                                    shape.Table.Rows[index].Cells[2].Merge(shape.Table.Rows[index].Cells[3]);

                                                //volgorde voor het liturgiebord is
                                                //liederen
                                                //lezen
                                                //tekst
                                                if (!legeregel)
                                                {
                                                    legeregel = true;
                                                }
                                                else if (!lezengehad)
                                                {
                                                    shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.Text = "L ";
                                                    shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.Text += Lezen;
                                                    shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                                                    lezengehad = true;
                                                }
                                                else if (!tekstgehad)
                                                {
                                                    shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.Text = "T ";
                                                    shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.Text += Tekst;
                                                    shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                                                    tekstgehad = true;
                                                }
                                                else
                                                {
                                                    shape.Table.Rows[index].Delete();
                                                    index--;
                                                }
                                            }
                                        }
                                        else
                                            shape.Table.Rows[index].Cells[1].Shape.TextFrame.TextRange.Text = "Liturgie";
                                    }
                                }
                            }
                        }
                    }
                    //voeg de slides in in het grote geheel
                    VoegSlideinPresentatiein(presentatie.Slides);
                    //sluit de geopende presentatie weer af
                    presentatie.Close();
                }
                #endregion Slides
                hoofdformulier.progressBar1.PerformStep();
                counterliturgieregel++;
            }

            //maximaliseer de presentatie ter controle voor de gebruiker
            objApp.WindowState = PpWindowState.ppWindowMaximized;

            hoofdformulier.autoEvent.Set();
        }

        /// <summary>
        /// Voeg een slide in in de hoofdpresentatie op de volgende positie (hoofdpresentatie werd aangemaakt bij het maken van deze klasse)
        /// </summary>
        /// <param name="slides">de slide die ingevoegd moet worden (voorwaarde is hierbij dat de presentatie waarvan de slide onderdeel is nog wel geopend is)</param>
        private void VoegSlideinPresentatiein(Slides slides)
        {
            foreach (Slide slide in slides)
            {
                //dit gedeelte is om het probleem van de eerste slide die al bestaat op te lossen voor alle andere gevallen maken we gewoon een nieuwe slide aan
                Slide voeginslide;
                if (slideteller == 1)
                    voeginslide = objPres.Slides[slideteller];
                else
                    voeginslide = objPres.Slides.AddSlide(slideteller, layout);

                //verwijder alle standaard toegevoegde dingen
                while (voeginslide.Shapes.Count > 0)
                {
                    voeginslide.Shapes[1].Delete();
                }
                //voeg de dingen van de template toe
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    shape.Copy();
                    voeginslide.Shapes.Paste();
                }

                slideteller++;
            }
        }
        #endregion
    }
}
