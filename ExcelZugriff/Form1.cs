//

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Globalization; // Monatsname
using System.IO; // speichern
using ExcelZugriff.Klassen;

namespace ExcelZugriff
{



    public partial class Form1 : Form
    {
        cExcelDB db;
        cKurzformen k;

        //cExcelConst ecConst;

        private List<string> _liste = new List<string>();




        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //

        }


        private TextBox iniTextBox(Control parent)
        {
            TextBox tb = new TextBox { };

            tb.Name = "tbTest";
            tb.Location = new Point(0, 10);
            tb.Text = "";
            tb.Dock = DockStyle.Fill;
            tb.Multiline = true;
            tb.Parent = parent;

            return tb;
        }


        private DataGrid iniDataGrid(Control parent)
        {
            DataGrid o = new DataGrid { };

            o.Name = "tbTest";
            o.ReadOnly = true;
            o.Location = new Point(0, 10);
            o.Text = "";
            o.Dock = DockStyle.Fill;
            //o.Multiline = true;
            
            o.Parent = parent;

            return o;
        }

        public string List2Str(List<string> l)
        {
            string str = "";
            foreach (string s in l) { str += s + Environment.NewLine; }
            return str;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Test(textBox1.Text);
        }

        public void Test2()
        {
            string fn = @"z:\2009.xlsx";

            db = new cExcelDB(fn);
            db.Connect();

            textBox2.Text = "aaa" + db.GetTableNames();


        }

        private void Test(string fn)
        {
            
            k = new cKurzformen();
            db = new cExcelDB(fn);


            // Form1.Text = fn;

            db.Connect();
            //textBox2.Text =  db.GetExcelData("März");

            List<string> l = db.GetTableNamesList();

            //   _liste.Clear();
            CreatecbWorksheets(l);
            BuildPages();


            //textBox2.Text = List2Str(_liste);

            //s = db.GetTableNames();





        }

        private void button2_Click(object sender, EventArgs e)
        {

        }



        public void BuildPages()
        {
            textBox2.Text = "";
            //List<TextBox> ltb = new List<TextBox>();
            List<DataGrid> ldg = new List<DataGrid>();

            tcWorksheets.TabPages.Clear();
            k.Clear();
            for (int i = 0; i < cbWorksheets.Items.Count; i++)
            {
                if (cbWorksheets.GetItemChecked(i) == true)
                {
                    string wsName = cbWorksheets.Items[i].ToString();
                    tcWorksheets.TabPages.Add(wsName);
                    ldg.Add(iniDataGrid(tcWorksheets.TabPages[(tcWorksheets.TabPages.Count - 1)]));
                    Kurzform(db.GetExcelData(wsName), _liste);

                    ldg[ldg.Count - 1].DataSource = db.GetExcelData(wsName);
                }
            }
            textBox2.Text = k.Show();
        }


        public void CreatecbWorksheets(List<string> namesOfWorksheets)
        {

            TextBox[] atb = new TextBox[] { };
            List<TextBox> ltb = new List<TextBox>();

            CultureInfo ci = new CultureInfo("de-DE");

            ltb.Clear();
            tcWorksheets.TabPages.Clear();
            cbWorksheets.Items.Clear();

            for (int monthID = 0; monthID < 12; monthID++)
            {
                foreach (string wsName in namesOfWorksheets)
                    if (wsName == ci.DateTimeFormat.MonthNames[monthID])
                    {
                        cbWorksheets.Items.Add(wsName);
                        cbWorksheets.SetItemChecked(cbWorksheets.Items.Count - 1, true);

                        tcWorksheets.TabPages.Add(wsName);
   

                        namesOfWorksheets.Remove(wsName);
                        break;
                    }
            }

            namesOfWorksheets.Sort();
            foreach (string wsName in namesOfWorksheets) { cbWorksheets.Items.Add(wsName); }
        }


        public List<string> Kurzform(DataTable dt, List<string> liste)
        {
            string s = "";
            DataRow drDatum = null;

            if (dt.Rows.Count > 0) { dt.Rows[0].Delete(); dt.AcceptChanges(); }

            if (dt.Rows.Count > 0)
            {
                drDatum = dt.Rows[0];

                //System.Type.GetType(dt.Rows[0]);
                //    drDatum.Field.
                //    dt.Rows[0].Delete(); dt.AcceptChanges();
            }
            foreach (DataRow row in dt.Rows)
                for (int i = 1; i < row.ItemArray.Length; i++)
                {
                    s = row[i].ToString();
                    if ((s != "") && (s.Length <= 6))
                    {
                        liste.Add(s);
                        string datum;
                        datum = drDatum[i].ToString();
                        //datum = drDatum[i].Field<DateTime>(

                        //   if ( drDatum[i].GetType() == Type.GetType(String) ) { datum += "$"; }
                        //  datum = drDatum[i].GetType().ToString();

                        this.k.Add(datum, s);
                        //myListe.Add(
                    }
                }
            return liste;
        }

        private void Form1_Leave(object sender, EventArgs e)
        {
        }

        private void cbWorksheets_KeyDown(object sender, KeyEventArgs e)
        {
            //            textBox1.Text =            e.KeyData.ToString();

            if (e.KeyData == (Keys.Control | Keys.A))
            {
                for (int i = 0; i < cbWorksheets.Items.Count; i++) cbWorksheets.SetItemChecked(i, true);
            }
        }

        

        private void miErzeugenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            BuildPages();

        }

        private void miSpeichernUnterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Stream myStream;

          //  SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string sFileName = saveFileDialog1.FileName;
                if (sFileName != "")
                {
                    string s = textBox2.Text;
                    try
                    {
                        using (StreamWriter outfile =  new StreamWriter(sFileName))
                        {
                            outfile.Write(s);
                        }
                    }
                    catch (Exception exc) { textBox2.Text = exc.Message + Environment.NewLine; }
                
                    
                }
                
                    
            }


        }

        private void miLadenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        //    OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Filter = "Excel 2010 files (*.xlsx)|*.xlsx|Excel (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (( openFileDialog1.OpenFile()) != null)
                {
                    string fn = openFileDialog1.FileName;
                    Test(fn);
                    textBox1.Text = fn;
                    
                
               
                }
            }

        }

        private void miHilfeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fn = @"z:\2009.xlsx";

            db = new cExcelDB(fn);
            textBox2.Text = db.Test();

        }

        private void miÜberToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fn = @"z:\2009.xlsx";

            db = new cExcelDB(fn);
            
            DataTable dt = db.TestDt();

            DataGrid dataGridView1 =  new DataGrid(); 

            dataGridView1.DataSource = dt;

            //tb.Name = "tbTest";
            dataGridView1.Location = new Point(0, 10);
            dataGridView1.Dock = DockStyle.Fill;
            //tb.Multiline = true;
            dataGridView1.Parent = panel3;


        }



    }
}
