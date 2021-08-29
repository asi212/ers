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
            string str3;
            string str4;
            string str5;
            string str6;
            string str7;
            string str8;

            int rCnt;
            int rows;
            int start;
            //  MessageBox.Show("testa");
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"X:\ERSTools\EndtestData\error_list.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //get first worksheet
                                                                              // MessageBox.Show("testb");
            range = xlWorkSheet.UsedRange;
            start = 1;
            rows = range.Rows.Count;
            string[,] LF_array = new string[rows + 1, 8];
           //MessageBox.Show();
            for (rCnt = start; rCnt <= start + rows - 1; rCnt++)
            {
               // MessageBox.Show(rCnt.ToString());
                str1 = (range.Cells[rCnt, 1] as Excel.Range).Text;// gets column 1(A) of error_list
                LF_array[rCnt - start, 0] = str1;
                str2 = (range.Cells[rCnt, 2] as Excel.Range).Text;// gets column 2(B) of error_list
                LF_array[rCnt - start, 1] = str2;
                str3 = (range.Cells[rCnt, 3] as Excel.Range).Text;// gets column 3(C) of error_list
                LF_array[rCnt - start, 2] = str3;
                str4 = (range.Cells[rCnt, 4] as Excel.Range).Text;// gets column 4(D) of error_list
                LF_array[rCnt - start, 3] = str4;
                str5 = (range.Cells[rCnt, 5] as Excel.Range).Text;// gets column 5(E) of error_list
                LF_array[rCnt - start, 4] = str5;
                str6 = (range.Cells[rCnt, 6] as Excel.Range).Text;// gets column 6(F) of error_list
                LF_array[rCnt - start, 5] = str6;
                str7 = (range.Cells[rCnt, 7] as Excel.Range).Text;// gets column 7(G) of error_list
                LF_array[rCnt - start, 6] = str7;
                str8 = (range.Cells[rCnt, 8] as Excel.Range).Text;// gets column 7(G) of error_list
                LF_array[rCnt - start, 7] = str8;
            }

            xlWorkBook.Close(false, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            update_errorsheet(LF_array, rows); // call next methode
        }

 

        private void update_errorsheet(string[,] LF_array, int rows)
        {

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            Excel.Range colRange;

            int last_row;

            //  MessageBox.Show("testa");
            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Open(@"X:\ERSTools\EndtestData\Production_errors.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); //get first worksheet
                                                                              // MessageBox.Show("testb");
            range = xlWorkSheet.UsedRange;
            last_row = range.Rows.Count; // num rows / last row

           // MessageBox.Show(range.Address);
           // MessageBox.Show(last_row.ToString());

            colRange = xlWorkSheet.Columns["A:A"];//get the range object where you want to search from


            string entry_number;
            int i;
            int k;
            k = 1;


            for (i = 0; i <= rows; i++) // lets loop through our array and see if each lf is already contained in the serial number spreadsheet
            {
                entry_number = LF_array[i, 0];
                Excel.Range resultRange = colRange.Find(What: entry_number, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext);// search lf in the range, if find result, return a range

                if (resultRange is null)
                {
                    xlWorkSheet.Cells[last_row + k, 1] = LF_array[i, 0]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 2] = LF_array[i, 1]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 3] = LF_array[i, 2]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 4] = LF_array[i, 3]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 5] = LF_array[i, 4]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 6] = LF_array[i, 5]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 7] = LF_array[i, 6]; // fill in first empty column A cell
                    xlWorkSheet.Cells[last_row + k, 8] = LF_array[i, 7]; // fill in first empty column A cell


                    k = k + 1;                                          // add one so the next thing we add is also to an empty cell
                    //MessageBox.Show(LF_array[i, 0] + " added to row " + (last_row + k).ToString());
                }
                else
                {
                    // Do nothing
                    // MessageBox.Show("found " + i.ToString());
                }

            }

            range = xlWorkSheet.UsedRange;
            last_row = range.Rows.Count; // num rows / last row
            // close seriernnummern XLS
            xlWorkBook.Save();
            xlWorkBook.Close(false);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);


        }
    }
}
