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

                try
                {
                    Workbook excelBook = excelApp.Workbooks.Open(filesArray[file]);
                    _Worksheet excelSheet = excelBook.Sheets[1];
                    Range excelRange = excelSheet.UsedRange;

                    //Forward ExcelBook
                    string sheetCellValue = excelSheet.Cells[2, 2].value;
                    if (sheetCellValue == "102" || sheetCellValue == "101")
                    {
                        excelSheet.Name = "Receipt";
                        //ShipmentsReceits(ref excelSheet);
                    }
                    else if (sheetCellValue == "602" || sheetCellValue == "601")
                    {
                        excelSheet.Name = "Shipments";
                        //ShipmentsReceits(ref excelSheet);
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

                        if (excelSheet.Cells[2, 4].value.Contains("HSE"))
                        {
                            SplitAndCountStock openStockClass = new SplitAndCountStock();
                            excelSheet.Name = "Current HSE3 stock";
                            openStockClass.Stock(ref excelSheet, ref excelBook);

                        }
                    }
                    
                    // This block should be removed or moved somewhere else
                    excelBook.SaveAs(@"C:\Users\ssladmin\Desktop\Weekly rep\HSE 2 Invoice.xlsx");
                    excelBook.Close(true);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                catch (Exception ex)
                {
                    Form1 msg = new Form1();
                    msg.sendMessage("Error occurred " + ex);
                }

                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    // Notify User that Reports were finished
                    Form1 msg = new Form1();
                    msg.sendMessage("Finished!");
                }
            }
        }
    }
}
