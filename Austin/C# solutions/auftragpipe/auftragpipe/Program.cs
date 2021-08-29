using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace auftragpipe
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Program n = new Program();
            n.read_excel();
        }

        private void read_excel()
        {
           
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str1;
            string str2;
            int rCnt;
            int rows;
            int start;
          //  MessageBox.Show("testa");
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"X:\Austin\Forecast\Auftrag.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //get first worksheet
           // MessageBox.Show("testb");
            range = xlWorkSheet.UsedRange;
            start = 2200;
            rows = 5000;
            string[,] LF_array = new string[rows+1,2];
           // MessageBox.Show("testc");
            for (rCnt = start; rCnt <= start+rows; rCnt++) 
            {
                str1 = (range.Cells[rCnt, 2] as Excel.Range).Text;// gets column 2(B) of Auftragslist
                LF_array[rCnt - start,0] = str1;
                str2 = (range.Cells[rCnt, 3] as Excel.Range).Text;// gets column 3(C) of Auftragslist
                LF_array[rCnt - start,1] = str2;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            // MessageBox.Show("test0");

            write_excel(LF_array, rows); // call next methode
        }

        private void write_excel(string[,] LF_array, int rows)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            // MessageBox.Show("test1");
            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            //  MessageBox.Show("test2");
            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            //  MessageBox.Show("test3");

            oSheet.get_Range("A1", "B" + (rows + 1).ToString()).Value = LF_array;


            oXL.Visible = false;
            oXL.UserControl = false;
            //  MessageBox.Show("test4");

            File.Delete("X:\\Austin\\Forecast\\pipe_backup.xlsx");
            System.IO.File.Copy("X:\\Austin\\Forecast\\pipe.xlsx", "X:\\Austin\\Forecast\\pipe_backup.xlsx", true);
            File.Delete("X:\\Austin\\Forecast\\pipe.xlsx");
            oWB.SaveAs("X:\\Austin\\Forecast\\pipe.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // MessageBox.Show("test5");
            //oWB.Close();

        }


        private void update_snxls()
        {

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str1;
            string str2;
            int rCnt;
            int rows;
            int start;
            //  MessageBox.Show("testa");
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"X:\Austin\Forecast\Seriennummern.xlsm", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //get first worksheet
                                                                              // MessageBox.Show("testb");
            range = xlWorkSheet.UsedRange;
            start = 2200;
            rows = 5000;
            string[,] LF_array = new string[rows + 1, 2];
            // MessageBox.Show("testc");
            for (rCnt = start; rCnt <= start + rows; rCnt++)
            {
                str1 = (range.Cells[rCnt, 2] as Excel.Range).Text;// gets column 2(B) of Auftragslist
                LF_array[rCnt - start, 0] = str1;
                str2 = (range.Cells[rCnt, 3] as Excel.Range).Text;// gets column 3(C) of Auftragslist
                LF_array[rCnt - start, 1] = str2;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);


        }
    }
}
