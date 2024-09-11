using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;

namespace CsvToSqlServer
{
    public class CsvToList
    {
        public List<List<string>> ReadCsvToList(string csvFilePath)
        {
            List<List<string>> data = new List<List<string>>();

            using (var reader = new StreamReader(csvFilePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                while (csv.Read())
                {
                    List<string> row = new List<string>();
                    for (int i = 0; csv.TryGetField<string>(i, out string field); i++)
                    {
                        row.Add(field);
                    }
                    data.Add(row);
                }
            }

            return data;
        }
    }
}

    
        