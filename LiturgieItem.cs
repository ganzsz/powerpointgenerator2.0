using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace PowerpointGenerater2
{
    class LiturgieItem
    {
        public string Titel = "";
        public string Aansluitend = "";

        public bool isValid = true;

        public bool bordje = false;
        public string[] bordregel = new string[3];
        private Form1 papa;
        public bool isLied = false;
        public bool eenvers = false;

        public string psalmmap;
        public List<int> verzen = new List<int>();
        public string presentatiepad = "";

        public LiturgieItem(string regel, Form1 pa)
        {
            this.papa = pa;
            List<string> onderdelen = new List<string>();
            char[] sep = { ' ', ':' };
            foreach (string s in regel.Split(sep, 3))
            {
                if (!s.Equals(""))
                    onderdelen.Add(s);
            }
            #region Zang
            if (onderdelen.Count()>1)
            {
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
                                    if (File.Exists(psalmmap + @"\" + i + @".txt") || File.Exists(psalmmap + @"\" + i + @".gif"))
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
                    this.bordregel[1] = onderdelen[1] + ":";
                    this.bordregel[2] = onderdelen[2];
                }
                else
                {
                    this.bordregel[1] = onderdelen[1];
                    this.bordregel[2] = "";
                }

                this.Aansluitend="Aansluitend: "+papa.instellingen.getMask(onderdelen[0]) + " " + onderdelen[1];
                if (onderdelen.Count() > 2)
                    this.Aansluitend += ": " + bordregel[2];
                //this.Aansluitend = "Aansluitend: " + papa.instellingen.getMask(onderdelen[0]) + " " + onderdelen[1] + ": ";
                this.Titel = papa.instellingen.getMask(onderdelen[0]) + " " + onderdelen[1];
                if(!eenvers)
                    this.Titel += ": ";
            }
            else
            #endregion
            {
                System.Diagnostics.Debug.WriteLine(papa.instellingen.Databasepad + @"\" + regel + @".pptx");
                if (File.Exists(papa.instellingen.Databasepad + @"\" + regel + @".pptx"))
                {
                    this.presentatiepad = papa.instellingen.Databasepad + @"\" + regel.ToLower() + @".pptx";
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
    }
}
