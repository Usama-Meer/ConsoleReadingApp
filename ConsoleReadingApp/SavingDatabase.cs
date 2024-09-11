

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace ConsoleReadingApp
{
    public class DatabaseSaver
    {
        private string connectionString;

        public DatabaseSaver(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public void SaveDataToDatabase(List<List<string>> data)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Create table if it doesn't exist
                string createTableQuery = @"
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'SalesData')
                BEGIN
                    CREATE TABLE SalesData (
                        OrderDate NVARCHAR(100),
                        Region NVARCHAR(50),
                        Manager NVARCHAR(50),
                        SalesMan NVARCHAR(50),
                        Item NVARCHAR(50),
                        Units INT,
                        Unit_price DECIMAL(10, 2),
                        Sale_amt DECIMAL(15, 2)
                    )
                END";
                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                {
                    createTableCommand.ExecuteNonQuery();
                }

                foreach (var row in data.Skip(0))
                {
                    if (row.Count == 8) // Ensure the row has exactly 8 columns
                    {
                        string query = "INSERT INTO SalesData (OrderDate, Region, Manager, SalesMan, Item, Units, Unit_price, Sale_amt) VALUES (@OrderDate, @Region, @Manager, @SalesMan, @Item, @Units, @Unit_price, @Sale_amt)";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@OrderDate", string.IsNullOrEmpty(row[0]) ? (object)DBNull.Value : row[0]);
                            command.Parameters.AddWithValue("@Region", string.IsNullOrEmpty(row[1]) ? (object)DBNull.Value : row[1]);
                            command.Parameters.AddWithValue("@Manager", string.IsNullOrEmpty(row[2]) ? (object)DBNull.Value : row[2]);
                            command.Parameters.AddWithValue("@SalesMan", string.IsNullOrEmpty(row[3]) ? (object)DBNull.Value : row[3]);
                            command.Parameters.AddWithValue("@Item", string.IsNullOrEmpty(row[4]) ? (object)DBNull.Value : row[4]);

                            if (int.TryParse(row[5], out int units))
                            {
                                command.Parameters.AddWithValue("@Units", units);
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@Units", DBNull.Value);
                            }

                            if (decimal.TryParse(row[6], out decimal unitPrice))
                            {
                                command.Parameters.AddWithValue("@Unit_price", unitPrice);
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@Unit_price", DBNull.Value);
                            }

                            if (decimal.TryParse(row[7], out decimal saleAmt))
                            {
                                command.Parameters.AddWithValue("@Sale_amt", saleAmt);
                            }
                            else
                            {
                                command.Parameters.AddWithValue("@Sale_amt", DBNull.Value);
                            }

                            command.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        // Handle the case where the row does not have the expected number of columns
                        Console.WriteLine("Skipping row due to incorrect number of columns.");
                    }
                }
            }
        }
    }
}