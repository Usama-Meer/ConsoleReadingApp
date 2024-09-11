using System;
using System.Collections.Generic;
using System.IO;

namespace ConsoleReadingApp
{
    public class CsvWriter
    {
        public void WriteToCsv(List<List<string>> data, string filePath)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    foreach (var row in data)
                    {
                        string line = string.Join(",", row);
                        writer.WriteLine(line);
                    }
                }
                Console.WriteLine($"Data successfully written to {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while writing to CSV: {ex.Message}");
            }
        }
    }
}