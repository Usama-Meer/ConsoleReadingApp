using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleReadingApp
{
    public class ExcelCreator
    {
        private Excel.Application excelApp;
        private Excel.Workbook excelBook;
        private Excel._Worksheet excelSheet;

        public ExcelCreator()
        {
            excelApp = new Excel.Application();
            if (excelApp == null)
            {
                throw new Exception("Excel is not installed!!");
            }
            excelBook = excelApp.Workbooks.Add();
            excelSheet = (Excel._Worksheet)excelBook.Sheets[1];
        }

        //public void WriteData()
        //{
        //    // Writing data to cells
        //    excelSheet.Cells[1, 1] = "Hello";
        //    excelSheet.Cells[1, 2] = "World!";
        //    excelSheet.Cells[2, 1] = "This is";
        //    excelSheet.Cells[2, 2] = "Interop";
        //}

        public void WriteData(List<List<string>> data)
        {
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    excelSheet.Cells[i + 1, j + 1] = data[i][j];
                }
            }
        }


        public void SaveAndClose(string filePath)
        {
            // Save the workbook
            excelBook.SaveAs(filePath);

            // Close the workbook and quit the application
            excelBook.Close();
            excelApp.Quit();

            // Release COM objects
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}