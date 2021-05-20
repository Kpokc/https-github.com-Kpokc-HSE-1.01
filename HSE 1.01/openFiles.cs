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

                Workbook excelBook = excelApp.Workbooks.Open(filesArray[file]);
                _Worksheet excelSheet = excelBook.Sheets[1];
                Range excelRange = excelSheet.UsedRange;

                //
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
                else if (excelSheet.Cells[2, 4].value == "HSE3")
                {
                    SplitAndCountStock openStockClass = new SplitAndCountStock();
                    excelSheet.Name = "Current HSE3 stock";
                    openStockClass.Stock(ref excelSheet, ref excelBook);

                }

                excelBook.SaveAs(@"C:\Users\ssladmin\Desktop\Weekly rep\HSE 2 Invoice.xlsx");
                excelBook.Close(true);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
