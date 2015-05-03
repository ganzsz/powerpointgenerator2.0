using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace PowerpointGenerater2
{
    public class Instellingen
    {
        public string Databasepad;
        public string Templateliederen;
        public string Templatetheme;
        public int regelsperslide = 6;
        private List<Mapmask> lijstmasks = new List<Mapmask>();
        public string TemplateAbeeldingLied;
        public bool dubbelePuntKolom = false;
        public string maskPath;

        public Instellingen()
            : this("", "", "", 6)
        {

        }

        public Instellingen(string databasepad, string templateliederen, string templatetheme, int regelsperslide)
        {
            this.Databasepad = databasepad;
            this.Templateliederen = templateliederen;
            this.Templatetheme = templatetheme;
            this.regelsperslide = regelsperslide;
        }

        public bool AddMask(Mapmask mask)
        {
            if (!lijstmasks.Contains(mask))
            {
                lijstmasks.Add(mask);
                return true;
            }
            return false;
        }

        public List<Mapmask> GetMasks()
        {
            return lijstmasks;
        }
        public void ClearMasks()
        {
            lijstmasks.Clear();
        }

        public static bool WriteXML(Instellingen instellingen, string path)
        {
            try
            {
                XmlWriterSettings xws = new XmlWriterSettings();
                xws.Indent = true;

                //schrijf instellingen weg
                XmlWriter xw = XmlWriter.Create(path + "instellingen.xml", xws);
                xw.WriteStartDocument();
                xw.WriteStartElement("Instellingen");
                xw.WriteStartElement("Databasepad");
                xw.WriteString(instellingen.Databasepad);
                xw.WriteEndElement();
                xw.WriteStartElement("Templateliederen");
                xw.WriteString(instellingen.Templateliederen);
                xw.WriteEndElement();
                xw.WriteStartElement("Templatetheme");
                xw.WriteString(instellingen.Templatetheme);
                xw.WriteEndElement();
                xw.WriteStartElement("RegelsperSlide");
                xw.WriteString(instellingen.regelsperslide.ToString());
                xw.WriteEndElement();
                xw.WriteStartElement("AfbeeldingLied");
                xw.WriteString(instellingen.TemplateAbeeldingLied.ToString());
                xw.WriteEndElement();
                xw.WriteStartElement("dubbelePuntKolom");
                xw.WriteString((instellingen.dubbelePuntKolom ? "1" : "0"));
                xw.WriteEndElement();
                xw.WriteStartElement("MaskPath");
                xw.WriteString(instellingen.maskPath);
                xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteEndDocument();

                xw.Flush();
                xw.Close();

                //schrijf Masks weg
                xw = XmlWriter.Create(path + "masks.xml", xws);
                xw.WriteStartDocument();
                xw.WriteStartElement("Masks");
                foreach (Mapmask mask in instellingen.lijstmasks)
                {
                    xw.WriteStartElement("Mask");
                    xw.WriteStartElement("Name");
                    xw.WriteString(mask.Name);
                    xw.WriteEndElement();
                    xw.WriteStartElement("RealName");
                    xw.WriteString(mask.RealName);
                    xw.WriteEndElement();
                    xw.WriteEndElement();
                }
                xw.WriteEndElement();
                xw.WriteEndDocument();

                xw.Flush();
                xw.Close();

                return true;
            }
            catch (PathTooLongException)
            {
                return false;
            }
        }

        public static Instellingen LoadXML(string path)
        {
            Instellingen instellingen = new Instellingen();
            XmlDocument xdoc = new XmlDocument();

            xdoc.Load(path + "instellingen.xml");
            XmlElement root = xdoc.DocumentElement;
            XmlNodeList nodelist = root.GetElementsByTagName("Databasepad");
            foreach (XmlNode item in nodelist)
                instellingen.Databasepad = item.InnerText;
            nodelist = root.GetElementsByTagName("Templateliederen");
            foreach (XmlNode item in nodelist)
                instellingen.Templateliederen = item.InnerText;
            nodelist = root.GetElementsByTagName("Templatetheme");
            foreach (XmlNode item in nodelist)
                instellingen.Templatetheme = item.InnerText;
            nodelist = root.GetElementsByTagName("RegelsperSlide");
            foreach (XmlNode item in nodelist)
            {
                bool result = System.Int32.TryParse(item.InnerText, out instellingen.regelsperslide);
                if (!result)
                    instellingen.regelsperslide = 6;
            }
            nodelist = root.GetElementsByTagName("AfbeeldingLied");
            foreach (XmlNode item in nodelist)
                instellingen.TemplateAbeeldingLied = item.InnerText;
            nodelist = root.GetElementsByTagName("dubbelePuntKolom");
            foreach (XmlNode item in nodelist)
                instellingen.dubbelePuntKolom = item.InnerText == "1";
            nodelist = root.GetElementsByTagName("MaskPath");
            foreach (XmlNode item in nodelist)
                instellingen.maskPath = item.InnerText;

            if (File.Exists(instellingen.maskPath))
                xdoc.Load(instellingen.maskPath);
            else
            {
                xdoc.Load(path + "masks.xml");
                instellingen.maskPath = path + "masks.xml";
            }
            root = xdoc.DocumentElement;
            nodelist = root.GetElementsByTagName("Name");
            XmlNodeList nodelist2 = root.GetElementsByTagName("RealName");
            if (nodelist.Count == nodelist2.Count)
            {
                for (int i = 0; i < nodelist.Count; i++)
                {
                    instellingen.lijstmasks.Add(new Mapmask(nodelist[i].InnerText, nodelist2[i].InnerText));
                }
            }
            return instellingen;
        }

        public override string ToString()
        {
            return string.Format("databasepad: {0}\n templateliederen: {1}\n templatetheme: {2}\n regels per slide: {3}\n", Databasepad, Templateliederen, Templatetheme, regelsperslide);
        }

        internal string getMask(string p)
        {
            foreach (Mapmask m in this.lijstmasks)
            {
                if (m.RealName.ToLower() == p.ToLower())
                    return m.Name;
            }
            return p;
        }


    }
}
