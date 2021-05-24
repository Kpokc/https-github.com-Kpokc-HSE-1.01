using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace HSE_1._01
{
    class SplitAndCountStock
    {
        Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        public void Stock(ref _Worksheet excelSheet, ref Workbook excelBook)
        {
            Form1 msg = new Form1();
            msg.sendMessage("Start Counting");

            Range excelRange = excelSheet.UsedRange;

            // Used range forward into array object
            object[,] values = (object[,])excelRange.Value2;


            List<string> eaList = new List<string>();
            List<string> ctList = new List<string>();
            List<string> haemList = new List<string>();
            List<string> ndList = new List<string>();
            List<string> pcccList = new List<string>();

            // Loop throughout "Material" column
            for (int i = 2; i <= values.Length / 15; i++)
            {
                if (values[i, 1].ToString().Contains("AE CARDS"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        eaList.Add(Convert.ToString(values[i, y + 1]));
                    }
                }

                if (values[i, 1].ToString().Contains("CLINICAL"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        ctList.Add(Convert.ToString(values[i, y + 1]));
                    }
                }

                if (values[i, 1].ToString().Contains("HAEM"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        haemList.Add(Convert.ToString(values[i, y + 1]));
                    }
                }

                if (values[i, 1].ToString().Contains("N+D"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        ndList.Add(Convert.ToString(values[i, y + 1]));
                    }
                }

                if (values[i, 1].ToString().Contains("PCCC"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        pcccList.Add(Convert.ToString(values[i, y + 1]));
                    }
                }
            }

            createSheet(eaList, "EA Sheet", ref excelBook);
            createSheet(ctList, "CLINICAL", ref excelBook);
            createSheet(haemList, "HAEM", ref excelBook);
            createSheet(ndList, "N+D", ref excelBook);
            createSheet(pcccList, "PCCC", ref excelBook);

            //msg.sendMessage("N&D " + cabinetCount.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy h:mm tt"));

        }

        /*static List<string> retlist() {
            List<string> lName = new List<string>();
            return lName;
        }*/


        static void createSheet(List<string> arrayList, string sheetName, ref Workbook excelBook) 
        {
            //Form1 msg = new Form1();
            Worksheet newWorksheet;
            newWorksheet = excelBook.Worksheets.Add();
            newWorksheet.Name = sheetName;

            newWorksheet.Columns["A:A"].ColumnWidth = 20;
            newWorksheet.Columns["B:E"].ColumnWidth = 5;
            newWorksheet.Columns["F:G"].ColumnWidth = 17;
            newWorksheet.Columns["H:H"].ColumnWidth = 5;
            newWorksheet.Columns["I:J"].ColumnWidth = 15;
            newWorksheet.Columns["K:L"].ColumnWidth = 10;
            newWorksheet.Columns["M:O"].ColumnWidth = 17;

            string[,] arr = new string[arrayList.Count / 15, 15];
            int row = 0;
            int col = 0;
            for (int i = 0; i < arrayList.Count; i++)
            {
                if (col == 15)
                {
                    row++;
                    col = 0;
                }
                arr[row, col] = arrayList[i];
                col++;
            }
            string range1 = "A1";
            string range2 = "O1";
            if (arrayList.Count != 0) {
                range2 = "O" + (arrayList.Count / 15).ToString();
            }

            //msg.sendMessage("1 = " + range1 + " 2 = " + range2);
            newWorksheet.Range[range1, range2].Value2 = arr;


       
            //newWorksheet.Columns["C:C"].NumberFormat = "###";
            /*formatRange.NumberFormat = "#,###,###";
            xlWorkSheet.Cells[1, 1] = "1234567890";*/
        }
    }
}