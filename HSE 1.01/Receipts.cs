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
    class Receipts
    {
        Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        public void Receipt(ref _Worksheet excelSheet, ref Workbook excelBook)
        {
            // User notification
            Form1 msg = new Form1();
            msg.sendMessage("Start Counting Receipts?");

            Range excelRange = excelSheet.UsedRange;

            /// Remove bg-color, add borders ////
            excelRange.Interior.ColorIndex = 0;
            excelRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            // Some styling was applied
            excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, 13]].Font.Bold = true;
            excelSheet.Columns["A:A"].ColumnWidth = 11;
            excelSheet.Columns["B:D"].ColumnWidth = 7;
            excelSheet.Columns["E:E"].ColumnWidth = 40;
            excelSheet.Columns["F:F"].ColumnWidth = 20;
            excelSheet.Columns["G:G"].ColumnWidth = 40;
            excelSheet.Columns["H:H"].ColumnWidth = 12;
            excelSheet.Columns["I:I"].ColumnWidth = 25;
            excelSheet.Columns["J:L"].ColumnWidth = 11;
            excelSheet.Columns["M:M"].ColumnWidth = 4;
            excelSheet.Columns["O:O"].ColumnWidth = 13;

            // Gets the Calendar instance associated with a CultureInfo.
            CultureInfo myCI = new CultureInfo("en-US");
            Calendar myCal = myCI.Calendar;

            // Gets the DTFI properties required by GetWeekOfYear.
            CalendarWeekRule myCWR = myCI.DateTimeFormat.CalendarWeekRule;
            DayOfWeek myFirstDOW = myCI.DateTimeFormat.FirstDayOfWeek;

            /// Below block gets week number from the first row ///
            string excelDate = excelSheet.Cells[2, 11].value.ToString();
            var tempDate = DateTime.Parse(excelDate);
            var tempWeek = myCal.GetWeekOfYear(tempDate, myCWR, myFirstDOW) - 1;

            // Next block Color lines by week number (two colors applied to separate weeks)
            string reversal;
            var reversalExists = false; // used later to send notification


            //BELOW is used only for PCCC invoicing in future could be re-factored into class

            // PCCC Department Lists and array
            string[,] pcccLocationArray = {
                { "Shantalla", "0" },
                { "Athenry", "0" },
                { "Tuam", "0" },
                { "Loughrea", "0" },
                { "Doughiska", "0" },
                { "Mountbellew", "0" },
                { "Ballinasloe", "0" },
                { "Mervue", "0" }
            };

            // Block that colors receipt sheet
            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {
                /// Below block gets week number of each next "row" ///
                excelDate = excelSheet.Cells[i, 11].value.ToString();
                // parse it to temporary date
                tempDate = DateTime.Parse(excelDate);
                // Finds to which week number of the year "tempDate" belongs
                var tempWeek2 = myCal.GetWeekOfYear(tempDate, myCWR, myFirstDOW) - 1;

                /// compares temporary week (first row) to each next rows week///
                if (tempWeek == tempWeek2)
                {
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 19;
                }
                else
                {
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 20;
                    tempWeek = tempWeek2 + 1;
                }
                /// Check if reversal done on SAP - color in red////
                reversal = excelSheet.Cells[i, 2].Value;
                reversal.ToString();
                if ((reversal == "102") || (reversal == "602"))
                {
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 46;
                    reversalExists = true;
                }

                // Notify User if any reversal found
                if (reversalExists)
                {
                    msg.sendMessage("Reversal found! ");
                    excelSheet.Cells[excelRange.Rows.Count + 2, 5].Value = "reversal done on SAP";
                    excelSheet.Cells[excelRange.Rows.Count + 2, 5].Interior.ColorIndex = 46;
                }
            }

            // Block that inserts line between receipts
            var tempNum = excelSheet.Cells[2, 1].value;
            int rowCount = excelRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                /// compares temporary receipt number / shipment to each next rows///
                if (tempNum != excelSheet.Cells[i, 1].value)
                {
                    // Document header text value to upper
                    var cellValue = excelSheet.Cells[i - 1, 9].value.ToUpper();
                    // Check if cell value contains any of dictionary locations
                    for (int l = 0; l < pcccLocationArray.Length / 2; l++)
                    {
                        // Add one to each location found in cell
                        if (cellValue.Contains(pcccLocationArray[l, 0].ToUpper()))
                        {
                            int num = Int16.Parse(pcccLocationArray[l, 1]);
                            num++;
                            pcccLocationArray[l, 1] = num.ToString();
                        }
                    }

                    tempNum = excelSheet.Cells[i, 1].value;
                    Range line = (Range)excelSheet.Rows[i];
                    line.Insert();
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 15;
                    /// ads one extra row to row count (rowNum) due to blank row added ///
                    rowCount += 1;
                }

                // When gets to last row (as receipt/pick can take only one row)
                if (i == rowCount)
                {
                    // Document header text value to upper
                    var cellValue = excelSheet.Cells[i - 1, 9].value.ToUpper();
                    // Check if cell value contains any of dictionary locations
                    for (int l = 0; l < pcccLocationArray.Length / 2; l++)
                    {
                        // Add one to each location found in cell
                        if (cellValue.Contains(pcccLocationArray[l, 0].ToUpper()))
                        {
                            int num = Int16.Parse(pcccLocationArray[l, 1]);
                            num++;
                            pcccLocationArray[l, 1] = num.ToString();
                        }
                    }
                }
            }
        }
    }
}
