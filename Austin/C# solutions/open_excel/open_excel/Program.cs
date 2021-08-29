using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace open_excel
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

            string str;
            int rCnt;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\a.ibele\Documents\Seriennummern_V2.xlsm", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            range = xlWorkSheet.UsedRange;
            string[] LF_array = new string[10000];
            string[] SN_array = new string[10000];

            for (rCnt = 3300; rCnt <= 9999; rCnt++) // column 1 fills in LF
            {       
                str = (range.Cells[rCnt, 1] as Excel.Range).Text; 
                LF_array[rCnt-1] = str;
            }
            for (rCnt = 3300; rCnt <= 9999; rCnt++) // column 5 fills in SN
            {
                str = (range.Cells[rCnt, 5] as Excel.Range).Text; 
                SN_array[rCnt - 1] = str;
            }


            string LF_num = "70000";
            // Check if selected LF value is in array //

            int pos = Array.IndexOf(LF_array, LF_num);
            if (pos > -1) // if the array contains LF_num
            {
                MessageBox.Show("exists");
            }
            else
            {
                MessageBox.Show("doesn't exists");
            }


            MessageBox.Show(LF_array.Length.ToString());

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }


    }
}
