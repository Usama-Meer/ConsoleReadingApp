

using System;
using System.Collections.Generic;
using ConsoleReadingApp;
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {


        //Reading Excel file

        string filePathRead = @"C:\Users\User\source\repos\ConsoleReadingApp\ConsoleReadingApp\SaleData.xlsx";

        //creating application
        FileReader fileReader = new FileReader(filePathRead);

        //reading data 

        fileReader.ReadData();


        var fileData= fileReader.CollectingData();
        fileReader.Close();
        Console.WriteLine("Excel data read successfully!");



        //Creating Excel file
        ExcelCreator excelCreator = new ExcelCreator();
        excelCreator.WriteData(fileData);
        string filePathWrite = @"C:\Users\User\Documents\MyExcel.xlsx";
        excelCreator.SaveAndClose(filePathWrite);
        Console.WriteLine("Excel file created and data written successfully!");


        //saving data inside the database
        string connectionString = "Server=DESKTOP-PADUH07\\MSSQLSERVER01; Database=ExcelDB; Trusted_Connection=true; TrustServerCertificate=true";
        DatabaseSaver databaseSaver = new DatabaseSaver(connectionString);
        databaseSaver.SaveDataToDatabase(fileData);


        SqlToExcel databaseReader = new SqlToExcel(connectionString);

        List<List<string>> dataFromDb = databaseReader.CollectingData();

        //Creating Excel file using data from database
        
        ExcelCreator sqlExcelCreator = new ExcelCreator();
        sqlExcelCreator.WriteData(dataFromDb);
        string sqlFilePathWrite = @"C:\Users\User\Documents\SqlExcel.xlsx";
        sqlExcelCreator.SaveAndClose(sqlFilePathWrite);
        Console.WriteLine("Excel file created and data written successfully!");



        Console.WriteLine("Excel data read and saved to database successfully!");
    }
}
