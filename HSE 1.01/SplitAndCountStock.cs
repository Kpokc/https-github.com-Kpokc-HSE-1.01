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
    public class Users
    {
        public int id = 0;
        public string name = string.Empty;
        public Users()
        {
            // Constructor Statements
        }
        public void GetUserDetails(int uid, string uname)
        {
            id = uid;
            uname = name;
            Console.WriteLine("Id: {0}, Name: {1}", id, name);
        }
        public int Designation { get; set; }
        public string Location { get; set; }
    }


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

            // Department Lists
            List<string> eaList = new List<string>();
            List<string> ctList = new List<string>();
            List<string> haemList = new List<string>();
            List<string> ndList = new List<string>();
            List<string> pcccList = new List<string>();

            // Loop throughout "Material" column
            for (int i = 2; i <= values.Length / 15; i++)
            {
                if (values[i, 1].ToString().Contains("AE CARDS") || values[i, 1].ToString().Contains("AE REGISTERS") || values[i, 1].ToString().Contains("AE FRACTURE"))
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

            // From List to sheet
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

            // Headers
            newWorksheet.Cells[1, 1].Value2 = "Material";
            newWorksheet.Cells[1, 3].Value2 = "Plnt";
            newWorksheet.Cells[1, 4].Value2 = "SLoc";
            newWorksheet.Cells[1, 5].Value2 = "S";
            newWorksheet.Cells[1, 6].Value2 = "Batch";
            newWorksheet.Cells[1, 7].Value2 = "Description";
            newWorksheet.Cells[1, 8].Value2 = "Typ";
            newWorksheet.Cells[1, 9].Value2 = "StorageBin";
            newWorksheet.Cells[1, 10].Value2 = "Available stock";
            newWorksheet.Cells[1, 11].Value2 = "BUn";
            newWorksheet.Cells[1, 12].Value2 = "GR Date";
            newWorksheet.Cells[1, 13].Value2 = "Pick quantity";
            newWorksheet.Cells[1, 14].Value2 = "Stock for putaway";
            newWorksheet.Cells[1, 15].Value2 = "Total Stock";
            newWorksheet.get_Range("A1", "O15").Font.Bold = true;

            // Column Width
            newWorksheet.Columns["A:A"].ColumnWidth = 20;
            newWorksheet.Columns["B:E"].ColumnWidth = 5;
            newWorksheet.Columns["F:G"].ColumnWidth = 17;
            newWorksheet.Columns["H:H"].ColumnWidth = 5;
            newWorksheet.Columns["I:J"].ColumnWidth = 15;
            newWorksheet.Columns["K:L"].ColumnWidth = 10;
            newWorksheet.Columns["M:O"].ColumnWidth = 17;

            // 2D array length and width
            string[,] arr = new string[arrayList.Count / 15, 15];

            int row = 0;
            int col = 0;

            for (int i = 0; i < arrayList.Count; i++)
            {
                // Next row
                if (col == 15)
                {
                    row++;
                    col = 0;
                }
                arr[row, col] = arrayList[i];
                col++;
            }

            // Get the Range where to load data 
            string range1 = "A2";
            string range2 = "O1";
            if (arrayList.Count != 0) {
                range2 = "O" + (arrayList.Count / 15).ToString();
            }

            //msg.sendMessage("1 = " + range1 + " 2 = " + range2);

            // Array to sheet
            newWorksheet.Range[range1, range2].Value2 = arr;

        }
    }
}