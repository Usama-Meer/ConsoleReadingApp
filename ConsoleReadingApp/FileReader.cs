using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ConsoleReadingApp
{
    public class FileReader
    {
        private Application excelApp;
        private Workbook excelBook;
        private _Worksheet excelSheet;
        private Range excelRange;

        public FileReader(string filePath)
        {
            excelApp = new Application();
            if (excelApp == null)
            {
                throw new Exception("Excel is not installed!!");
            }
            excelBook = excelApp.Workbooks.Open(filePath);
            excelSheet = (_Worksheet)excelBook.Sheets[1];
            excelRange = excelSheet.UsedRange;
        }

        public void ReadData()
        {
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        Console.Write(excelRange.Cells[i, j].Value.ToString() + "\t");
                }
                Console.WriteLine();
            }
        }

        public List<List<string>> CollectingData()
        {
            List<List<string>> data = new List<List<string>>();
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                List<string> row = new List<string>();
                for (int j = 1; j <= colCount; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {

                        row.Add(excelRange.Cells[i, j].Value.ToString());
                    }
                    else
                    {
                        row.Add(string.Empty);
                    }
                }
                data.Add(row);
            }

            return data;
        }


        public void Close()
        {
            excelBook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}
