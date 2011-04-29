using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelZugriff.Klassen
{
    class cExcelDB
    {
        private string _fileName = "";

        //   #region constructors
        public cExcelDB()
        {
        }


        public cExcelDB(string fileName)
        {
            if (File.Exists(fileName))
                _fileName = fileName;
        }


        //     #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// 
        public string FilePath
        {
            get { return _fileName; }
            set { _fileName = value; }
        }


        public OleDbConnection Connect()
        {
            //string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=Excel 8.0;";
            // imex=1 --- typ raten http://munishbansal.wordpress.com/2009/12/15/importing-data-from-excel-having-mixed-data-types-in-a-column-ssis/ 
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Password=\"\";User ID=Admin;Data Source=" + _fileName + ";Mode=Share Deny Write;Extended Properties=\"HDR=NO;\";Jet OLEDB:Engine Type=37;IMEX=1";
            OleDbConnection conn = null;
            if (_fileName != "") conn = new OleDbConnection(connStr);


            return conn;
        }

        public string GetTableNames()
        {
            string str = "";
            foreach (string t in GetTableNamesList()) { str += t + Environment.NewLine; }
            return str;
        }


        public List<string> GetTableNamesList()
        {
            List<string> tables = new List<string>();

            try
            {
                OleDbConnection conn = Connect();

                if (conn != null)
                {
                    conn.Open();

                    string s = "";
                    DataTable sheetTable = conn.GetSchema("Tables");
                    foreach (DataRow row in sheetTable.Rows)
                    {
                        s = row[2].ToString();
                        if (s[s.Length - 1] == '$')
                        { s = s.Remove(s.Length - 1); tables.Add(s); };
                    }

                    s = Show(sheetTable);
                    conn.Close();
                }
            }
            catch (Exception e) { string s = e.Message + Environment.NewLine; }
            return tables;
        }


        public DataTable GetExcelData(string wsName)
        {
            string s = "";
            DataTable dt = null;
            try
            {
                OleDbConnection conn = Connect();
                if (conn != null)
                {

                    OleDbDataAdapter DataAdapter = new OleDbDataAdapter("SELECT * FROM [" + wsName + "$]", conn);

                    dt = new DataTable(wsName);

                    DataAdapter.Fill(dt);
                    conn.Close();
                }
            }
            catch (Exception e) { s += e.Message + Environment.NewLine; }

            return dt;
        }


        public string GetExcelDataStr(string wsName)
        {
            DataTable dt = GetExcelData(wsName);
            return Show(dt);
        }


        public string Show(DataTable dt, string delimiter = "\t")
        {
            string s = "";


            //            s += dt.Rows.Count.ToString();
            //            s += Environment.NewLine;

            DataColumn dc;
            
            foreach (DataRow row in dt.Rows)
            {
                
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    //row[i]
                    s += row[i].ToString() + delimiter;
          /*
                    object val = row[i];
                    if (val is int) s += "{int}";
                    if (val is double) s += "{double}";
                    if (val is DateTime) s += "{datetime}";
                    if (val is string) s += "{string}";
                    */

                }
                s += Environment.NewLine;
            }

            return s;
        }

        public string Test()
        {
            return "";
        }

        public DataTable TestDt()
        {
            string s = "";
            /*
            OleDbConnection conn = Connect();

            OleDbDataAdapter DataAdapter = new OleDbDataAdapter("SELECT * FROM [" + "Tabelle1" + "$]", conn);
            DataTable dt = conn.GetSchema();

             * 
             */


            DataTable worksheets;

            _fileName = @"Z:\Mappe1.xls";
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Password=\"\";User ID=Admin;Data Source=" + _fileName + ";Mode=Share Deny Write;Extended Properties=\"HDR=NO;\";Jet OLEDB:Engine Type=37;IMEX=1";
     
            using (OleDbConnection connection =
                      new OleDbConnection(connStr))
            {
                connection.Open();
                worksheets = connection.GetSchema("Tables");
            }

            DataTable columns;
            
            string[] restrictions = { "Tabelle1$" };
            using (OleDbConnection connection =
                                 new OleDbConnection(connStr))
            {
                connection.Open();
 //               columns = connection.GetSchema("Columns", restrictions);
 
                columns = connection.GetSchema("Columns");
            }

            

       //     s = Show(worksheets);
            s += Show(columns);

            return columns;
        }


    }




}
