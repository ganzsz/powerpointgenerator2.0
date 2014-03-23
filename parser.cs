using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerpointGenerater2
{
   
    class parser
    {
        private Form1 papa;
        private List<string> liturgie;
        private List<LiturgieItem> items = new List<LiturgieItem>();
        public string errormsg = "";

        public parser(Form1 pa, string liturgie)
        {
            this.papa = pa;
            this.liturgie = new List<string>(liturgie.Split('\n'));
        }

        /// <summary>
        /// Function to parse the liturgie
        /// </summary>
        /// <returns>Returns false if one of the slides wasn't found, then use parser.errormsg</returns>
        public bool parse()
        {
            this.liturgie.Reverse(0,liturgie.Count);
            for (int i = 0; i < this.liturgie.Count; i++)
            {
                items.Add(new LiturgieItem(this.liturgie[i],papa));
            }
            return true;
        }
    }
}
