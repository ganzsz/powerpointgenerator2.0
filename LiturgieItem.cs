using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PowerpointGenerater2
{
    class LiturgieItem
    {
        public string Titel = "";
        public string Aansluitend = "";

        public bool isValid = false;

        public bool bordje = false;
        private Form1 papa;

        public LiturgieItem(string regel, Form1 pa)
        {
            this.papa=pa;
            List<string> onderdelen = new List<string>();
            char[] sep = {' ',':'};
            foreach (string s in regel.Split(sep))
            {
                if (!s.Equals(""))
                    onderdelen.Add(s);
            }
#region Zang
            if (Directory.Exists(papa.instellingen.Databasepad + @"\" + onderdelen[0].ToLower()))
            {
                if (Directory.Exists(papa.instellingen.Databasepad + @"\" + onderdelen[0].ToLower() + @"\" + onderdelen[1].ToLower()))
                {
                    if (onderdelen.Count > 2)
                    {
                        foreach (string vers in onderdelen[2].Split(','))
                        {

                        }
                    }
                    else
                    {

                    }
                }
            }
#endregion
        }
    }
}
