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

            Worksheet newWorksheet;

            // Used range forward into array object
            object[,] values = (object[,])excelRange.Value2;

            //msg.sendMessage(values[29, 15].ToString());

            List<string> eaList = new List<string>();
            eaList.Add("AE Stock");

            // AE array length is same as spreadsheet's (+1 as excel first cell starts from 1:1)
            /*string[,] aeArray = new string[,] { };
            msg.sendMessage((aeArray.Length).ToString());*/
            
            // Loop throughout "Material" column
            for (int i = 2; i <= values.Length/15; i++)
            {
                if (values[i, 1].ToString().Contains("AE CARDS"))
                {
                    for (int y = 0; y < 15; y++)
                    {
                        eaList.Add(Convert.ToString(values[i, y + 1]));

                        //msg.sendMessage(aeArray[i, y].ToString());
                    }

                }
            }
            msg.sendMessage("List count = " + (eaList.Count/16).ToString());

            newWorksheet = excelBook.Worksheets.Add();
            newWorksheet.Name = "AE Stock";
            int row = 1;
            int col = 1;
            for (int i = 1; i < eaList.Count; i++)
            {
                if (col == 16) {
                    row++;
                    col = 1;
                }
                newWorksheet.Cells[row, col].Value = eaList[i];
                col++;
            }

            var filesCount = 0;
            var cabinetCount = 0;
            int rowCount = excelRange.Rows.Count;

            

            

            /*for (int i = 2; i <= rowCount; i++)
            {
                row++;
                // Get Document header text
                var cellValue = excelSheet.Cells[i, 1].value.ToUpper();

                // EA Department
                if (cellValue.Contains("AE"))
                {
                    //string sheetName = "AE Stock";
                    //addSheetCopyToSheet(sheetName, sheetList, i, row, ref excelSheet, ref excelBook);
                    *//*// If Cards
                    if (cellValue.Contains("AE CARDS")) {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 15;
                    }
                    // If Registers
                    if (cellValue.Contains("AE REGISTERS"))
                    {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 36;
                    }*//*
                    filesCount++;
                }

                // Clinical Trials Department
                if (cellValue.Contains("CLINICAL TRIAL"))
                {
                    //string sheetName = "CLTrial Stock";
                    //addSheetCopyToSheet(sheetName, sheetList, i, row, ref excelSheet, ref excelBook);
                    *//*// If Box
                    if (excelSheet.Cells[i, 6].value.ToUpper() == "BOX") {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 4;
                    }
                    // If Envelope
                    else if (excelSheet.Cells[i, 6].value.ToUpper() == "ENVELOPE") {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 6;
                    }*//*
                }

                // HAEM Department
                if (cellValue.Contains("HAEM"))
                {
                    //string sheetName = "HAEM Stock";
                    //addSheetCopyToSheet(sheetName, sheetList, i, row, ref excelSheet, ref excelBook);
                    *//*// If HSE Box
                    if (cellValue.Contains("HSE_BOX")) {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 10;
                    }
                    // If SSL Box
                    if (cellValue.Contains("SSL_BOX"))
                    {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 12;
                    }*//*
                }

                // N&D Department
                if (cellValue.Contains("N+D BOX"))
                {
                    //string sheetName = "N&D Stock";
                    //addSheetCopyToSheet(sheetName, sheetList, i, row, ref excelSheet, ref excelBook);
                    *//*excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 24;
                    cabinetCount++;*//*
                }

                // PCCC Department
                if (cellValue.Contains("PCCC"))
                {
                    //string sheetName = "PCCC Stock";
                    //addSheetCopyToSheet(sheetName, sheetList, i, row, ref excelSheet, ref excelBook);
                    *//*// If HSE Box
                    if (cellValue.Contains("BOX") || cellValue.Contains("FILES"))
                    {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 19;
                    }
                    // If SSL Box
                    if (cellValue.Contains("CABINET"))
                    {
                        excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 20;
                    }*//*
                }
            }*/
            msg.sendMessage("AE :" + filesCount.ToString());
            msg.sendMessage("N&D " + cabinetCount.ToString() + " " + DateTime.Now.ToString("MM/dd/yyyy h:mm tt"));

        }


        /*static void addSheetCopyToSheet(string sheetName, List<string> sheetList, int i, int row, ref _Worksheet excelSheet, ref Workbook excelBook)
        {
            Form1 msg = new Form1();

            Range excelRange = excelSheet.UsedRange;
            Worksheet newWorksheet, destSheet;
            Range sourceRange, destRange, desRan;

            if (sheetList.Contains(sheetName))
            {
                /// Copy paste all report lines into correct tab ///
                newWorksheet = excelBook.Worksheets.get_Item(sheetName);
                Range xlRange = newWorksheet.UsedRange;
                int rowNum = xlRange.Rows.Count;

                sourceRange = excelSheet.Rows[i];
                destSheet = excelBook.Worksheets[sheetName];
                desRan = destSheet.UsedRange;
                destRange = destSheet.Rows[rowNum + 1];
                sourceRange.Copy(destRange);
            }
            else if (!sheetList.Contains(sheetName))
            {
                msg.sendMessage(sheetName.ToString());
                /// Creating new sheet by Vendors name ///
                sheetList.Add(sheetName);
                newWorksheet = excelBook.Worksheets.Add();
                newWorksheet.Name = sheetName;

                /// Copy paste first line with headers ///
                sourceRange = excelSheet.Rows[1];
                destSheet = excelBook.Worksheets[sheetName];
                destRange = destSheet.Rows[1];
                sourceRange.Copy(destRange);

                /// Copy paste first line of materials///
                sourceRange = excelSheet.Rows[i];
                destSheet = excelBook.Worksheets[sheetName];
                desRan = destSheet.UsedRange;
                row = desRan.Rows.Count;
                destRange = destSheet.Rows[2];
                sourceRange.Copy(destRange);
            }
        }*/
    }
}
