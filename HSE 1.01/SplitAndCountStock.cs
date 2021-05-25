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
    public class Stocks
    {
        // Auto-implemented properties.
        public string Description { get; set; }
        public string Unit { get; set; }
        public int Qty { get; set; }

        /*Stocks eaStock = new Stocks { Description = "EaStocks", Unit = "Box", Qty = 10 };
        Console.WriteLine(cat.Name);*/
        // Class - ----------------------------------Stock Name ---------- Unit ---Qty
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
            List<string> aeList = new List<string>();
            // Class - ----------------------------------Stock Name ---------- Unit ---Qty
            Stocks aeStock1 = new Stocks { Description = "AE Stocks", Unit = "Cards", Qty = 0 };
            Stocks aeStock2 = new Stocks { Description = "AE Stocks", Unit = "Registers", Qty = 0 };
            Stocks aeStock3 = new Stocks { Description = "AE Stocks", Unit = "Fracture", Qty = 0 };

            List<string> ctList = new List<string>();
            // Class - ----------------------------------Stock Name ---------- Unit ---Qty
            Stocks ctStock1 = new Stocks { Description = "CT Stocks", Unit = "Box", Qty = 0 };
            Stocks ctStock2 = new Stocks { Description = "CT Stocks", Unit = "Envelope", Qty = 0 };

            List<string> haemList = new List<string>();
            // Class - ----------------------------------Stock Name ---------- Unit ---Qty
            Stocks haemStock1 = new Stocks { Description = "HAEM Stocks", Unit = "HSE Box", Qty = 0 };
            Stocks haemStock2 = new Stocks { Description = "HAEM Stocks", Unit = "SSL Box", Qty = 0 };

            List<string> ndList = new List<string>();
            // Class - ----------------------------------Stock Name ---------- Unit ---Qty
            Stocks ndStock1 = new Stocks { Description = "N+D Stocks", Unit = "ND Box", Qty = 0 };

            List<string> pcccList = new List<string>();
            // Class - ----------------------------------Stock Name ---------- Unit ---Qty
            Stocks pcccStock1 = new Stocks { Description = "PCCC Stocks", Unit = "PCCC Box + Files", Qty = 0 };
            Stocks pcccStock2 = new Stocks { Description = "PCCC Stocks", Unit = "PCCC Cabinet", Qty = 0 };

            // Loop throughout "Material" column
            for (int i = 1; i <= values.Length / 15; i++)
            {
                //msg.sendMessage(values[i, 1].ToString());
                // AE Stock count
                if (values[i, 1].ToString().Contains("AE CARDS") || values[i, 1].ToString().Contains("AE REGISTERS") || values[i, 1].ToString().Contains("AE FRACTURE"))
                {
                    // All available stock to list
                    for (int y = 0; y < 15; y++)
                    {
                        aeList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    // Add to Cards Stock (Material Column)
                    if (values[i, 1].ToString().Contains("AE CARDS")){
                        aeStock1.Qty += Int16.Parse(values[i, 10].ToString());
                    }

                    // Add to Registers Stock (Material Column)
                    if (values[i, 1].ToString().Contains("AE REGISTERS"))
                    {
                        aeStock2.Qty += Int16.Parse(values[i, 10].ToString());
                    }

                    // Add to Fracture Stock (Material Column)
                    if (values[i, 1].ToString().Contains("AE FRACTURE"))
                    {
                        aeStock3.Qty += Int16.Parse(values[i, 10].ToString());
                    }
                }

                // Clinical Trials Stock count
                if (values[i, 1].ToString().Contains("CLINICAL TRIAL"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        ctList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    // Add to Box Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("BOX"))
                    {
                        ctStock1.Qty += Int16.Parse(values[i, 10].ToString());
                    }

                    // Add to Envelope Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("ENVELOPE"))
                    {
                        ctStock2.Qty += Int16.Parse(values[i, 10].ToString());
                    }
                }

                // HAEM Stock count
                if (values[i, 1].ToString().Contains("HAEM"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        haemList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    // Add to Fracture Stock (Material Column)
                    if (values[i, 1].ToString().Contains("HSE_BOX"))
                    {
                        haemStock1.Qty += Int16.Parse(values[i, 10].ToString());
                    }

                    // Add to Fracture Stock (Material Column)
                    if (values[i, 1].ToString().Contains("SSL_BOX"))
                    {
                        haemStock2.Qty += Int16.Parse(values[i, 10].ToString());
                    }
                }

                // N+D Stock count
                if (values[i, 1].ToString().Contains("N+D"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        ndList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    ndStock1.Qty += Int16.Parse(values[i, 10].ToString());
                }

                // PCCC Stock count
                if (values[i, 1].ToString().Contains("PCCC"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        pcccList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    // Add to Cards Stock (Material Column)
                    if (values[i, 1].ToString().Contains("BOX") || values[i, 1].ToString().Contains("FILES"))
                    {
                        pcccStock1.Qty += Int16.Parse(values[i, 10].ToString());
                    }

                    // Add to Cards Stock (Material Column)
                    if (values[i, 1].ToString().Contains("CABINET"))
                    {
                        pcccStock2.Qty += Int16.Parse(values[i, 10].ToString());
                    }
                }
            }

            // From List to sheet
            createSheet(aeList, "AE Stocks", ref excelBook);
            createSheet(ctList, "CT Stocks", ref excelBook);
            createSheet(haemList, "HAEM Stocks", ref excelBook);
            createSheet(ndList, "N+D Stocks", ref excelBook);
            createSheet(pcccList, "PCCC Stocks", ref excelBook);

        }

        /*static List<string> retlist() {
            List<string> lName = new List<string>();
            return lName;
        }*/

        static void createSheet(List<string> arrayList, string sheetName, ref Workbook excelBook) 
        {
            Form1 msg = new Form1();
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
            newWorksheet.get_Range("A1", "Q1").Font.Bold = true;

            // Column Width
            newWorksheet.Columns["A:A"].ColumnWidth = 20;
            newWorksheet.Columns["B:E"].ColumnWidth = 5;
            newWorksheet.Columns["F:G"].ColumnWidth = 17;
            newWorksheet.Columns["H:H"].ColumnWidth = 5;
            newWorksheet.Columns["I:J"].ColumnWidth = 15;
            newWorksheet.Columns["K:L"].ColumnWidth = 10;
            newWorksheet.Columns["M:O"].ColumnWidth = 17;

            // 2D array length and width
            string[,] arr = new string[((arrayList.Count) / 15), 15];

            int row = 0;
            int col = 0;

            for (int i = 0; i < arr.Length ; i ++) {
                arr[row,col] = arrayList[i];
                col++;
                if (col == 15) {
                    col = 0;
                    if (row != arr.Length / 15) {
                        row++;
                    }
                }
            }

            // Get the Range where to load data 
            string range1 = "A2";
            string range2 = "O1";
            if (arrayList.Count != 0) {
                range2 = "O" + ((arrayList.Count / 15)+1).ToString();
            }

            //msg.sendMessage("1 = " + range1 + " 2 = " + range2);

            // Array to sheet
            newWorksheet.Range[range1, range2].Value2 = arr;

        }
    }
}