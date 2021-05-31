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

            // PCCC Department Lists and array
            string[,] pcccLocationArray = {
                // LOCATION // NUMBER OF TRIPS // FILES RECEIVED // BOXES RECEIVED // CABINETS RECEIVED
                { "Shantalla", "0", "0", "0", "0" },
                { "Athenry", "0", "0", "0", "0" },
                { "Tuam", "0", "0", "0", "0" },
                { "Loughrea", "0", "0", "0", "0" },
                { "Doughiska", "0", "0", "0", "0" },
                { "Mountbellew", "0", "0", "0", "0" },
                { "Ballinasloe", "0", "0", "0", "0" },
                { "Mervue", "0", "0", "0", "0" }
            };
            /////////////////////////////////////////////////////////////////////////////////////
            
            // AE Department Lists and array
            string[,] aeArray = {
                // Unit // QTY
                { "Nr of Receipts", "0" },
                { "Cards", "0"},
                { "Register", "0"},
                { "Fracture", "0"}
            };
            /////////////////////////////////////////////////////////////////////////////////////

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
            var documentHeaderCell = "";
            var tempNum = excelSheet.Cells[2, 1].value;
            int rowCount = excelRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                /// compares temporary receipt number / shipment to each next rows///
                if (tempNum != excelSheet.Cells[i, 1].value)
                {

                    ////////////////////////////////////////   PCCC BLOCK   //////////////////////////////////////////////////////////
                    // Document header text value to upper
                    var cellValue = excelSheet.Cells[i, 9].value.ToUpper();

                    // Check if cell value contains any of the locations
                    for (int l = 0; l < 8; l++)
                    {
                        // Add one to each location if found in cell
                        if (cellValue.Contains(pcccLocationArray[l, 0].ToUpper()))
                        {
                            int num = Int16.Parse(pcccLocationArray[l, 1]);
                            num++;
                            pcccLocationArray[l, 1] = num.ToString();

                            // Add file qty to Location
                            if ((cellValue.Contains("FILE")))
                            {
                                documentHeaderCell = excelSheet.Cells[i, 9].value.ToUpper();
                                int findStr = documentHeaderCell.IndexOf("FILE");
                                string number = documentHeaderCell.Substring(0, findStr);
                                int fileNum = Int16.Parse(pcccLocationArray[l, 2]);
                                fileNum += Int32.Parse(number);
                                pcccLocationArray[l, 2] = fileNum.ToString();
                            }

                            // Add box qty to Location
                            if (cellValue.Contains("B_"))
                            {
                                documentHeaderCell = excelSheet.Cells[i, 9].value.ToUpper();
                                int findStr = documentHeaderCell.IndexOf("B_");
                                string number = documentHeaderCell.Substring(0, findStr);
                                int fileNum = Int16.Parse(pcccLocationArray[l, 3]);
                                fileNum += Int32.Parse(number);
                                pcccLocationArray[l, 3] = fileNum.ToString();
                            }

                            // Add Cabinets qty to Location
                            if (cellValue.Contains("CABINET"))
                            {
                                documentHeaderCell = excelSheet.Cells[i, 9].value.ToUpper();
                                int findStr = documentHeaderCell.IndexOf("CABINET");
                                string number = documentHeaderCell.Substring(0, findStr);
                                int fileNum = Int16.Parse(pcccLocationArray[l, 4]);
                                fileNum += Int32.Parse(number);
                                pcccLocationArray[l, 4] = fileNum.ToString();
                            }
                        }
                    }
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                    var materialCellValue = excelSheet.Cells[i, 6].value.ToUpper();

                    if (materialCellValue.ToString().Contains("AE CARDS") || materialCellValue.ToString().Contains("AE REGISTERS") || materialCellValue.ToString().Contains("AE FRACTURE"))
                    {
                        if (!cellValue.Contains("PART OF"))
                        {
                            int receiptsNum = Int16.Parse(aeArray[0, 1]);
                            receiptsNum++;
                            aeArray[0, 1] = receiptsNum.ToString();
                        
                            // Add to Cards Stock (Material Column)
                            if (materialCellValue.Contains("AE CARDS"))
                            {
                                documentHeaderCell = excelSheet.Cells[i, 9].value.ToUpper();
                                int findStr = documentHeaderCell.IndexOf("B_");
                                string number = documentHeaderCell.Substring(0, findStr);
                                int fileNum = Int16.Parse(aeArray[1, 1]);
                                fileNum += Int32.Parse(number);
                                aeArray[1, 1] = fileNum.ToString();
                            }

                            // Add to Registers Stock (Material Column)
                            if (materialCellValue.Contains("AE REGISTERS"))
                            {
                                documentHeaderCell = excelSheet.Cells[i, 9].value.ToUpper();
                                int findStr = documentHeaderCell.IndexOf("REGISTER");
                                string number = documentHeaderCell.Substring(0, findStr);
                                int fileNum = Int16.Parse(aeArray[2, 1]);
                                fileNum += Int32.Parse(number);
                                aeArray[2, 1] = fileNum.ToString();
                            }

                            // Add to Fracture Stock (Material Column)
                            if (materialCellValue.Contains("AE FRACTURE"))
                            {
                                documentHeaderCell = excelSheet.Cells[i, 9].value.ToUpper();
                                int findStr = documentHeaderCell.IndexOf("FRACTURE");
                                string number = documentHeaderCell.Substring(0, findStr);
                                int fileNum = Int16.Parse(aeArray[3, 1]);
                                fileNum += Int32.Parse(number);
                                aeArray[3, 1] = fileNum.ToString();

                            }
                        }
                    }

                    tempNum = excelSheet.Cells[i, 1].value;
                }

            }

            // Add lines between receipts
            var tNum = excelSheet.Cells[2, 1].value;
            int rCount = excelRange.Rows.Count;
            for (int i = 2; i <= rCount; i++)
            {
                if (tNum != excelSheet.Cells[i, 1].value)
                {
                    tNum = excelSheet.Cells[i, 1].value;
                    Range line = (Range)excelSheet.Rows[i];
                    line.Insert();
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 15;
                    /// ads one extra row to row count (rowNum) due to blank row added ///
                    rCount += 1;
                }
            }

            /// This block to remove with forwarding arrays to class
            Worksheet newWorksheet;
            newWorksheet = excelBook.Worksheets.Add();
            newWorksheet.Range["A1","E8" ].Value2 = pcccLocationArray;
            newWorksheet.Range["G1", "H4"].Value2 = aeArray;
       }
    }
}
