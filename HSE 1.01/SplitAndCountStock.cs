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
            msg.sendMessage("Start Counting?");

            Range excelRange = excelSheet.UsedRange;

            int rowCount = excelRange.Rows.Count;

            /*Range bRangeMulti = excelSheet.get_Range("A"+(rowCount - 5).ToString(), "A" + rowCount.ToString());
            bRangeMulti.Delete(XlDeleteShiftDirection.xlShiftUp);*/

            // Used range forward into array object
            object[,] values = (object[,])excelRange.Value2;

            // AE Department Lists and array
            List<string> aeList = new List<string>();
            string[,] aeArray = new string[4, 4];
            aeArray[0, 0] = "Cards";
            aeArray[0, 1] = "0";
            aeArray[1, 0] = "Registers";
            aeArray[1, 1] = "0";
            aeArray[2, 0] = "Fracture";
            aeArray[2, 1] = "0";
            aeArray[3, 0] = "";
            aeArray[3, 1] = "";

            // CT Department Lists and array
            List<string> ctList = new List<string>();
            string[,] ctArray = new string[4, 4];
            ctArray[0, 0] = "Box";
            ctArray[0, 1] = "0";
            ctArray[1, 0] = "Envelope";
            ctArray[1, 1] = "0";
            ctArray[2, 0] = "";
            ctArray[2, 1] = "";
            ctArray[3, 0] = "";
            ctArray[3, 1] = "";

            // HAEM Department Lists and array
            List<string> haemList = new List<string>();
            string[,] haemArray = new string[4, 4];
            haemArray[0, 0] = "HSE Box";
            haemArray[0, 1] = "0";
            haemArray[1, 0] = "SSL Box";
            haemArray[1, 1] = "0";
            haemArray[2, 0] = "";
            haemArray[2, 1] = "";
            haemArray[3, 0] = "";
            haemArray[3, 1] = "";

            // N+D Department Lists and array
            List<string> ndList = new List<string>();
            string[,] ndArray = new string[4, 4];
            ndArray[0, 0] = "ND Box";
            ndArray[0, 1] = "0";
            ndArray[1, 0] = "";
            ndArray[1, 1] = "";
            ndArray[2, 0] = "";
            ndArray[2, 1] = "";
            ndArray[3, 0] = "";
            ndArray[3, 1] = "";

            // PCCC Department Lists and array
            List<string> pcccList = new List<string>();
            string[,] pcccArray = new string[4, 4];
            pcccArray[0, 0] = "PCCC Box + Files";
            pcccArray[0, 1] = "0";
            pcccArray[1, 0] = "PCCC Cabinet";
            pcccArray[1, 1] = "0";
            pcccArray[2, 0] = "";
            pcccArray[2, 1] = "";
            pcccArray[3, 0] = "";
            pcccArray[3, 1] = "";

            List<string> bloodList = new List<string>();
            string[,] bloodArray = new string[4, 4];
            bloodArray[0, 0] = "18L + File Box";
            bloodArray[0, 1] = "0";
            bloodArray[1, 0] = "42L";
            bloodArray[1, 1] = "0";
            bloodArray[2, 0] = "64L";
            bloodArray[2, 1] = "0";
            bloodArray[3, 0] = "84L";
            bloodArray[3, 1] = "0";

            // Loop throughout "Material" column "-5" as last 5 rows are counts
            for (int i = 1; i <= (values.Length / 15); i++)
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
                        int num = Int16.Parse(aeArray[0, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        aeArray[0, 1] = num.ToString();

                    }

                    // Add to Registers Stock (Material Column)
                    if (values[i, 1].ToString().Contains("AE REGISTERS"))
                    {
                        int num = Int16.Parse(aeArray[1, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        aeArray[1, 1] = num.ToString();

                    }

                    // Add to Fracture Stock (Material Column)
                    if (values[i, 1].ToString().Contains("AE FRACTURE"))
                    {
                        int num = Int16.Parse(aeArray[2, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        aeArray[2, 1] = num.ToString();

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
                        int num = Int16.Parse(ctArray[0, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        ctArray[0, 1] = num.ToString();
                    }

                    // Add to Envelope Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("ENVELOPE"))
                    {
                        int num = Int16.Parse(ctArray[1, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        ctArray[1, 1] = num.ToString();
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
                        int num = Int16.Parse(haemArray[0, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        haemArray[0, 1] = num.ToString();
                    }

                    // Add to Fracture Stock (Material Column)
                    if (values[i, 1].ToString().Contains("SSL_BOX"))
                    {
                        int num = Int16.Parse(haemArray[1, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        haemArray[1, 1] = num.ToString();
                    }
                }

                // N+D Stock count
                if (values[i, 1].ToString().Contains("N+D"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        ndList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    int num = Int16.Parse(ndArray[0, 1]);
                    num += Int16.Parse(values[i, 10].ToString());
                    ndArray[0, 1] = num.ToString();
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
                        int num = Int16.Parse(pcccArray[0, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        pcccArray[0, 1] = num.ToString();
                    }

                    // Add to Cards Stock (Material Column)
                    if (values[i, 1].ToString().Contains("CABINET"))
                    {
                        int num = Int16.Parse(pcccArray[1, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        pcccArray[1, 1] = num.ToString();
                    }
                }

                // BloodBank Stock count
                if (values[i, 1].ToString().Contains("BLOODBANK"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        bloodList.Add(Convert.ToString(values[i, y + 1]));
                    }

                    // Add to 18L + File Box Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("18L") || values[i, 6].ToString().Contains("FILE BOX"))
                    {
                        int num = Int16.Parse(bloodArray[0, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        bloodArray[0, 1] = num.ToString();
                    }

                    // Add to 42L Box Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("42L"))
                    {
                        int num = Int16.Parse(bloodArray[1, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        bloodArray[1, 1] = num.ToString();
                    }

                    // Add to 64 Box Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("64L"))
                    {
                        int num = Int16.Parse(bloodArray[2, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        bloodArray[2, 1] = num.ToString();
                    }

                    // Add to 84 Box Stock (Batch Column)
                    if (values[i, 6].ToString().Contains("84L"))
                    {
                        int num = Int16.Parse(bloodArray[3, 1]);
                        num += Int16.Parse(values[i, 10].ToString());
                        bloodArray[3, 1] = num.ToString();
                    }
                }
            }

            // From List to sheet
            //msg.sendMessage(pcccList.Count.ToString());
            createSheet(aeList, "AE Stocks", aeArray, ref excelBook);
            createSheet(ctList, "CT Stocks", ctArray, ref excelBook);
            createSheet(haemList, "HAEM Stocks", haemArray, ref excelBook);
            createSheet(ndList, "N+D Stocks", ndArray, ref excelBook);
            createSheet(pcccList, "PCCC Stocks", pcccArray, ref excelBook);
            createSheet(bloodList, "Blood Stocks", bloodArray, ref excelBook);

        }

        static void createSheet(List<string> arrayList, string sheetName, string[,] array_count, ref Workbook excelBook) 
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
            newWorksheet.Columns["Q:Q"].ColumnWidth = 16;

            // Align Center
            newWorksheet.Columns["C:C"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            newWorksheet.Columns["H:H"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            newWorksheet.Columns["J:J"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            newWorksheet.Columns["M:O"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            newWorksheet.Columns["R:R"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

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
            // Array to sheet
            newWorksheet.Range[range1, range2].Value2 = arr;
            // Borders from top left cell to bottom right
            Range excelRange = newWorksheet.Range["A1", range2];
            excelRange.Borders.LineStyle = XlLineStyle.xlContinuous;


            //Heading of counted stock
            newWorksheet.Cells[1, 17].Value2 = sheetName;
            // Counted stock
            string rangeA = "Q2";
            string rangeB = "R5";
            newWorksheet.Range[rangeA, rangeB].Value2 = array_count;
            Range excelRangeSmall = newWorksheet.Range["Q1", "R5"];
            excelRangeSmall.Borders.LineStyle = XlLineStyle.xlContinuous;

        }
    }
}