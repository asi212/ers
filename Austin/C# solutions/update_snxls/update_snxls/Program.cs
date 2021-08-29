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
            n.update_snxls();
        }

        private void update_snxls()
        {
            // declare pipe.xls variables 
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str1;
            string str2;
            string IAsubstr = "IA";
            int rCnt;
            int rows;
            int start;
            int j;


            //  MessageBox.Show("testa");
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"X:\Austin\Forecast\pipe.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //get first worksheet

            Excel.Range last = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            range = xlWorkSheet.UsedRange;

            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;


            start = 1;
            rows = lastUsedRow;
            string[,] update_array = new string[rows, 2];
            j = 0;
            for (rCnt = start; rCnt <= start + rows; rCnt++)
            {
                str1 = (range.Cells[rCnt, 1] as Excel.Range).Text;// gets column 1(A) LF #, of intmediate sheet
                if (!str1.Contains(IAsubstr)) // if the LF is not an IA LF
                {
                    str2 = (range.Cells[rCnt, 1] as Excel.Range).Text;// gets column 2(B) ERS #, of intmediate sheet
                    update_array[j, 0] = str1;
                    update_array[j, 1] = str2; // fill in array with details of LF's that do not contain IA
                    j = j + 1;
                }
            }

            // close pipe XLS after we are done putting relevent data into our update_array
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);


            // declare seriern nummern xls variables
            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range colRange;
            string path = @"X:\Austin\Forecast\Seriennummern.xlsm";

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open(path, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(2); //get second worksheet

            Excel.Range start2 = xlWorkSheet2.Range["A1"];
            Excel.Range bottom2 = xlWorkSheet2.Range["A" + (xlWorkSheet2.UsedRange.Rows.Count + 1)];
            Excel.Range end2 = bottom2.End[Excel.XlDirection.xlUp];
            Excel.Range column2 = xlWorkSheet2.Range[start2, end2];

            string lastcellA = end2.Address.ToString();
            lastcellA = lastcellA.Substring(3); // remove first three characters
            int lastUsedRow2 = Int32.Parse(lastcellA);

            colRange = xlWorkSheet2.Columns["A:A"];//get the range object where you want to search from


            string lf;
            int i;
            int k;
            k = 1;
            for (i = 0; i <= j; i++) // lets loop through our array and see if each lf is already contained in the serial number spreadsheet
            {
                lf = update_array[i, 0];
                Excel.Range resultRange = colRange.Find(What: lf, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext);// search lf in the range, if find result, return a range

                if (resultRange is null)
                {
                    xlWorkSheet2.Cells[lastUsedRow2 + k, 1] = update_array[i, 0];
                   // MessageBox.Show((lastUsedRow2 + k).ToString());
                    k = k + 1;
                }
                else
                {
                   // MessageBox.Show("LF found");
                }

            }

            // close seriernnummern XLS
            xlWorkBook2.Save();
            xlWorkBook2.Close();
            xlApp2.Quit();

            Marshal.ReleaseComObject(xlWorkSheet2);
            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);

        }
    }
}
