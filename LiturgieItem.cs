using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace PowerpointGenerater2
{
    class LiturgieItem
    {
        public string Titel = "";
        public string Aansluitend = "";

        public bool isValid = true;

        public bool bordje = false;
        public string[] bordregel = new string[4];
        private Form1 papa;
        public bool isLied = false;
        public bool eenvers = false;

        public string psalmmap;
        public List<int> verzen = new List<int>();
        public string presentatiepad = "";

        private string tos = "NULL";

        public LiturgieItem(string regel, Form1 pa)
        {
            this.papa = pa;
            this.tos = regel;
            List<string> onderdelen = new List<string>();
            string[] verstemp = regel.Split(':');
            if (verstemp.Count() > 1)
            {
                if (verstemp[0].Split(' ').Count() > 1)
                {
                    onderdelen.Add(verstemp[0].Split(' ')[0]);
                    onderdelen.Add(verstemp[0].Split(' ')[1]);
                    onderdelen.Add(verstemp[1]);
                }
                else
                {
                    this.isValid = false;
                }
            }
            else
            {
                char[] sep = {' '};
                foreach (string s in regel.Split(sep, 2))
                {
                    if (!s.Equals(""))
                        onderdelen.Add(s);
                }
            }
            /*
             * Als er verzen zijn, spaties wegwerken ivm werking
             */
            if (onderdelen.Count > 2)
            {
                onderdelen[2] = Regex.Replace(onderdelen[2], " +", "", RegexOptions.Compiled);
            }
            #region Zang
            if (onderdelen.Count() > 1 && onderdelen[0].ToLower() != "lezen" && onderdelen[0].ToLower() != "tekst")
            {
                string regelt = onderdelen[0].ToLower();
                switch (regelt)
                {
                    case "ps":
                        onderdelen[0] = "psalm";
                        break;
                    case "gz":
                        onderdelen[0] = "gezang";
                        break;
                    case "opw":
                        onderdelen[0] = "opwekking";
                        break;
                    case "ld":
                    case "lb":
                        onderdelen[0] = "lied";
                        break;
                }
                string psalmmap = papa.instellingen.Databasepad + @"\" + onderdelen[0] + @"\" + onderdelen[1].ToLower();
                if (Directory.Exists(psalmmap))
                {
                    this.psalmmap = psalmmap;
                    this.isLied = true;
                    if (onderdelen.Count > 2)
                    {
                        foreach (string vers in onderdelen[2].Split(','))
                        {
                            if (vers.Split('-').Count() > 1)
                            {
                                for (int i = Int32.Parse(vers.Split('-')[0]); i <= Int32.Parse(vers.Split('-')[vers.Split('-').Count() - 1]); i++)
                                {
                                    if (File.Exists(psalmmap + @"\" + i + @".txt") || File.Exists(psalmmap + @"\" + i + @"-1.gif"))
                                        verzen.Add(i);
                                    else
                                        this.isValid = false;
                                }
                            }
                            else
                            {
                                if (File.Exists(psalmmap + @"\" + vers + @".txt") || File.Exists(psalmmap + @"\" + vers + @"-1.gif"))
                                    verzen.Add(Int32.Parse(vers));
                                else
                                    this.isValid = false;
                            }
                        }
                    }
                    else
                    {
                        for (int i = 1; i < 100; i++)
                        {
                            if (File.Exists(psalmmap + @"\" + i + @".txt") || File.Exists(psalmmap + @"\" + i + @"-1.gif"))
                                verzen.Add(i);
                            else
                                break;
                        }
                        if (verzen.Count == 1)
                            eenvers = true;
                    }
                }
                else
                {
                    this.isValid = false;
                }

                this.bordje = true;
                this.bordregel[0] = papa.instellingen.getMask(onderdelen[0]);
                if (onderdelen.Count() > 2)
                {
                    this.bordregel[1] = onderdelen[1];
                    this.bordregel[2] = ":";
                    this.bordregel[3] = onderdelen[2];
                    this.bordregel[3] = Regex.Replace(this.bordregel[3], ",+", ", ");
                }
                else
                {
                    this.bordregel[1] = onderdelen[1];
                    this.bordregel[2] = "";
                    this.bordregel[3] = "";
                }

                this.Aansluitend = "Aansluitend: " + papa.instellingen.getMask(onderdelen[0]) + " " + onderdelen[1];
                if (onderdelen.Count() > 2)
                    this.Aansluitend += ": " + bordregel[3];
                //this.Aansluitend = "Aansluitend: " + papa.instellingen.getMask(onderdelen[0]) + " " + onderdelen[1] + ": ";
                this.Titel = papa.instellingen.getMask(onderdelen[0]) + " " + onderdelen[1];
                if (!eenvers)
                    this.Titel += ": ";
            }
            else
            #endregion
            {
                System.Diagnostics.Debug.WriteLine(papa.instellingen.Databasepad + @"\" + onderdelen[0] + @".pptx");
                if (File.Exists(papa.instellingen.Databasepad + @"\" + onderdelen[0] + @".pptx"))
                {
                    this.presentatiepad = papa.instellingen.Databasepad + @"\" + onderdelen[0].ToLower() + @".pptx";
                    #region lezen
                    if (onderdelen[0].ToLower() == "lezen" || onderdelen[0].ToLower() == "tekst")
                    {
                        RichTextBox el;
                        if (onderdelen[0].ToLower() == "lezen")
                        {
                            el = papa.textBox1;
                        }
                        else
                        {
                            el = papa.textBox5;
                        }
                        if (onderdelen.Count() > 1)
                        {
                            int regelnummer = Int32.Parse(onderdelen[1]);

                            if (regelnummer <= el.Lines.Count() && regelnummer > 0)
                            {
                                this.Titel = el.Lines[regelnummer-1];
                                this.Aansluitend = "Aansluitend: "+papa.instellingen.getMask(onderdelen[0]) + ": " + this.Titel;
                            }
                            else
                            {
                                this.Titel = el.Text;
                                this.Aansluitend = "Aansluitend: " + papa.instellingen.getMask(onderdelen[0]);
                            }
                        }
                        else
                        {
                            this.Titel = el.Text;
                            this.Aansluitend = "Aansluitend: " + papa.instellingen.getMask(onderdelen[0]);
                        }
                    }
                    #endregion lezen;
                    else
                    {
                        if (!papa.instellingen.getMask(onderdelen[0]).Equals(""))
                            this.Aansluitend = "Aansluitend: " + papa.instellingen.getMask(onderdelen[0]);
                        else
                            this.Aansluitend = "";
                    }
                }
                else
                    this.isValid = false;

            }
            if (!this.isValid)
            {
                MessageBox.Show('"' + regel + "\" is niet gevonden.");
                return;
            }
        }

        public override string ToString()
        {
            return this.tos;
        }
    }
}
