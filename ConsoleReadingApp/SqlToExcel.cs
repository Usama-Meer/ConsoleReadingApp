using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ConsoleReadingApp
{


    public class SqlToExcel
    {
        private string connectionString;

        public SqlToExcel(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public List<List<string>> CollectingData()
        {
            List<List<string>> data = new List<List<string>>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT OrderDate, Region, Manager, SalesMan, Item, Units, Unit_price, Sale_amt FROM SalesData";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            List<string> row = new List<string>
                        {
                            reader["OrderDate"].ToString(),
                            reader["Region"].ToString(),
                            reader["Manager"].ToString(),
                            reader["SalesMan"].ToString(),
                            reader["Item"].ToString(),
                            reader["Units"].ToString(),
                            reader["Unit_price"].ToString(),
                            reader["Sale_amt"].ToString()
                        };
                            data.Add(row);
                        }
                    }
                }
            }

            return data;
        }
    }

}
