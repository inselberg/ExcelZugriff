using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;


namespace ExcelZugriff.Klassen
{
    class cKurzformen
    {

        // struct
        public struct MyTyp
        {
            public string name;
            public int zähler;
            public List<string> liste;// = new List<string>();

            public MyTyp(string name, int zähler, List<string> liste)
            {
                this.name = name;
                this.zähler = zähler;
                this.liste = liste;
            }

            public void inc() { this.zähler++; }
        }

        private List<string> idx = new List<string>();
        private List<MyTyp> myListe = new List<MyTyp>();

        // constructior
        public cKurzformen()
        {
        }

        public void Clear() { idx.Clear(); myListe.Clear(); }

        public void Add(string datum, string name)
        {
            name = name.ToLower();

            bool valid = true;
            foreach (char c in name)
            {
                if (c >= '0' && c <= '9') { valid = false; break; }
            }

            if (valid)
            {
                int i = idx.IndexOf(name);
                if (i == -1)
                {
                    idx.Add(name);
                    MyTyp t = new MyTyp(name, 1, new List<string>());
                    t.liste.Add(datum);
                    myListe.Add(t);
                }
                else
                {
                    MyTyp t = myListe[i];
                    t.inc();
                    t.liste.Add(datum);
                    myListe[i] = t;
                }
            }
        }

        public List<string> ToList()
        {
            List<string> l = new List<string>();
            string s = "";

            foreach (MyTyp t in myListe)
            {
                s = "";
                s += t.name + "\t";
                s += t.zähler.ToString() + "\t";
                foreach (string datum in t.liste) s += '[' + datum + ']' + " ";
                l.Add(s);
            }


            return l;
        }



        public string Show()
        {
            List<string> l = ToList();

            l.Sort();

            string str = "";
            foreach (string s in l)
            {
                str += s + Environment.NewLine;

            }
            return str;
        }
    }



}
