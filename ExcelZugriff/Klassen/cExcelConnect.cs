using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;


namespace ExcelZugriff.Klassen
{
    class cExcelConnect
    {

        string dummy = "mooo!";
        // 

        private string fileName;

        private Microsoft.Office.Interop.Excel.Application excel;
        private Workbook workbook;


        public List<string> msgs = new List<string>();
        public List<string> namesOfWorksheets = new List<string>();




        private void msg(string s) { msgs.Add(s); dummy += s + Environment.NewLine; }

        public cExcelConnect()
        {
            //excel = new Microsoft.Office.Interop.Excel.Application();
            //            wb = null;
            
        }


        public Workbook Open(string fileName)
        {
            this.fileName = fileName;

            // Excel starten
            excel = new Microsoft.Office.Interop.Excel.Application();
        //    excel.Visible = true;
            excel.Visible = false;

            Workbook wb = null;
            wb = excel.Workbooks.Open(
                 fileName,
                 ExcelConst.UpdateLinks.DontUpdate,
                 ExcelConst.ReadOnly,
                 ExcelConst.Format.Nothing,
                 "", // Passwort
                 "", // WriteResPasswort
                 ExcelConst.IgnoreReadOnlyRecommended,
                 XlPlatform.xlWindows,
                 "", // Trennzeichen

                 ExcelConst.NotEditable,
                //      ExcelConst.Editable,
           ExcelConst.Notify,
                //ExcelConst.DontNotifiy,
                 ExcelConst.Converter.Default,
                 ExcelConst.DontAddToMru,
                 ExcelConst.Local,
                 ExcelConst.CorruptLoad.NormalLoad);
            // Arbeitsblätter lesen

            workbook = wb;
            
            return wb;
        }

        public bool Close() { return Close(workbook); }

        public bool Close(Workbook wb)
        {
            wb.Close(false, null, null);
            excel.Quit();
            return true;
        }



        public bool LoadExcelFile_OK(string fileName)
        {

            namesOfWorksheets.Clear();
            //Open();


            Microsoft.Office.Interop.Excel.Application excel = null;
            Workbook wb = null;
            try
            {
                // Excel starten
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;

                // Datei öffnen

                wb = excel.Workbooks.Open(
                    fileName,
                    ExcelConst.UpdateLinks.DontUpdate,
                    ExcelConst.ReadOnly,
                    ExcelConst.Format.Nothing,
                    "", // Passwort
                    "", // WriteResPasswort
                    ExcelConst.IgnoreReadOnlyRecommended,
                    XlPlatform.xlWindows,
                    "", // Trennzeichen

                    ExcelConst.NotEditable,
                    //      ExcelConst.Editable,
              ExcelConst.Notify,
                    //ExcelConst.DontNotifiy,
                    ExcelConst.Converter.Default,
                    ExcelConst.DontAddToMru,
                    ExcelConst.Local,
                    ExcelConst.CorruptLoad.NormalLoad);
                // Arbeitsblätter lesen


                // 
                Sheets sheets = wb.Worksheets;
                

                //      msg("try2");
                //    msg("sheets.Count=" + sheets.Count);

                // namesOfWorksheets
                for (int i = 1; i <= sheets.Count; i++)
                    namesOfWorksheets.Add(((Worksheet)sheets[i]).Name);

                /*
                // ein Arbeitsblatt auswählen…
                Worksheet ws = (Worksheet)sheets.get_Item("Tabelle1");
                // …oder eine Zelle
                Range range = (Range)ws.get_Range("A1", "A1");
                // deren Wert auslesen
                string zellwert = range.Value2.ToString();
                Console.WriteLine(zellwert);
                 * 
                 */
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.Message);
                msg("e.Message:" + e.Message);
            }
            finally
            {
                wb.Close(false, null, null);
                excel.Quit();
            }

            return true;
        }

        //public bool Close() { wb  }

        public string GetData(string workSheetname)
        {



            string s = "muhh @" + workSheetname;
            /*
            Sheets sheets = workbook.Worksheets;
         
            Worksheet ws = (Worksheet)sheets.get_Item(workSheetname);

           // s += ws.Range.ToString();
            //ws.Columns();

            //Range range = (Range)ws.get_Range("A1", "A1");
            Range range = (Range)ws.UsedRange;

            
            int uc = ws.UsedRange.Count;
       //     int c = ws.Columns.Count;
            //int r = ws.Rows.Count;


            int c = range.Columns.Count;
            int r = range.Rows.Count;

        /*
            MsgBox "Anzahl der Spalten = " & oRgn.Columns.Count & vbCrLf & _
    "Anzahl der Zeilen = " & oRgn.Rows.Count & vbCrLf & _
    "Oberste Reihe = " & oCell.Row & vbCrLf & _
    "Äußerst linke Spalte = " & oCell.Column

            


            msg("ws.UsedRange.Count="+uc.ToString());
            msg("ws.Columns.Count=" + c.ToString());
            msg("ws.Rows.Count=" + r.ToString());

           //ws.UsedRange.Cells.get_Value

          //  Range v = (Range)ws.get_Range(1,1);
          //  string zellwert = v.Value2.ToString();
         //   msg("Zellwert: "+zellwert);

            //       ws.Delete();
            */
            return s;
        }

        public string GetDummy() { return dummy; }
        public string Test() { msg("Test()"); return dummy; }
        public void Clear() { dummy = ""; msgs.Clear(); }



















    }
}
