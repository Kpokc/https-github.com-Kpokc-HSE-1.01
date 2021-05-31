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
    class openFiles
    {
        Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        public void openFile(string[] filesArray) 
        {
            for (int file = 0; file < filesArray.Length; file++){

                /*Workbook excelBook = excelApp.Workbooks.Open(filesArray[file]);
                _Worksheet excelSheet = excelBook.Sheets[1];
                Range excelRange = excelSheet.UsedRange;

                //sendMsg.sendMessage(arr[arri]);

                string sheetCellValue = excelSheet.Cells[2, 2].value;
                if (sheetCellValue == "102" || sheetCellValue == "101")
                {
                    Receipts forward = new Receipts();
                    excelSheet.Name = "Receipt";
                    forward.Receipt(ref excelSheet, ref excelBook);
                }
                else if (sheetCellValue == "602" || sheetCellValue == "601")
                {
                    Shipments forward = new Shipments();
                    excelSheet.Name = "Shipments";
                    forward.Shipment(ref excelSheet, ref excelBook);
                }
                else
                {
                    // Delete two left columns
                    int colCount = excelRange.Columns.Count;
                    for (int twoLeftColumns = 1; twoLeftColumns <= 2; twoLeftColumns++)
                    {
                        Range column = (Range)excelSheet.Columns[1];
                        column.Delete();
                    }

                    // Delete top five rows from Backup
                    for (int fiveToprows = 1; fiveToprows <= 5; fiveToprows++)
                    {
                        Range line = (Range)excelSheet.Rows[1];
                        line.Delete();
                    }

                    // Delete bottom five rows from Backup
                    int rowCount = excelRange.Rows.Count;
                    for (int fiveToprows = 1; fiveToprows <= 5; fiveToprows++)
                    {
                        Range line = (Range)excelSheet.Rows[rowCount - 4];
                        line.Delete();
                    }

                    // Delete Second row
                    Range midLine = (Range)excelSheet.Rows[2];
                    midLine.Delete();

                    // Borders
                    excelRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                    excelSheet.Name = "Current stock";
                    SplitAndCountStock forward = new SplitAndCountStock();
                    forward.Stock(ref excelSheet, ref excelBook);
                    //Stock(ref excelSheet);

                }*/



                try
                {
                    Workbook excelBook = excelApp.Workbooks.Open(filesArray[file]);
                    _Worksheet excelSheet = excelBook.Sheets[1];
                    Range excelRange = excelSheet.UsedRange;

                    //Forward ExcelBook
                    string sheetCellValue = excelSheet.Cells[2, 2].value;
                    if (sheetCellValue == "102" || sheetCellValue == "101")
                    {
                        Receipts forward = new Receipts();
                        excelSheet.Name = "HSE Receipts";
                        forward.Receipt(ref excelSheet, ref excelBook);
                    }
                    else if (sheetCellValue == "602" || sheetCellValue == "601")
                    {
                        Shipments forward = new Shipments();
                        excelSheet.Name = "HSE Shipments";
                        forward.Shipment(ref excelSheet, ref excelBook);
                    }

                    // Temporary comment Stock section

                    else
                    {
                        // Delete two left columns
                        int colCount = excelRange.Columns.Count;
                        for (int twoLeftColumns = 1; twoLeftColumns <= 2; twoLeftColumns++)
                        {
                            Range column = (Range)excelSheet.Columns[1];
                            column.Delete();
                        }

                        // Delete top five rows from Backup
                        for (int fiveToprows = 1; fiveToprows <= 5; fiveToprows++)
                        {
                            Range line = (Range)excelSheet.Rows[1];
                            line.Delete();
                        }

                        // Delete bottom five rows from Backup
                        int rowCount = excelRange.Rows.Count;
                        for (int fiveToprows = 1; fiveToprows <= 5; fiveToprows++)
                        {
                            Range line = (Range)excelSheet.Rows[rowCount - 4];
                            line.Delete();
                        }

                        // Delete Second row
                        Range midLine = (Range)excelSheet.Rows[2];
                        midLine.Delete();

                        // Borders
                        excelRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                        if (excelSheet.Cells[2, 4].value.Contains("HSE"))
                        {
                            SplitAndCountStock forward = new SplitAndCountStock();
                            excelSheet.Name = "Current HSE3 stock";
                            forward.Stock(ref excelSheet, ref excelBook);

                        }
                    }
                }

                catch (Exception ex)
                {
                    Form1 msg = new Form1();
                    msg.sendMessage("Error occurred " + ex);
                }

                finally
                {
                    excelApp.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // Notify User that Reports were finished
                    Form1 msg = new Form1();
                    msg.sendMessage("Finished!");
                }
            }
            // This block should be removed or moved somewhere else
            
        }
    }
}
