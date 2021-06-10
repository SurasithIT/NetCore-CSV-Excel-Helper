using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;

namespace NetCore.CSV.Excel.Helper
{
    public class CSVandExcelHelper
    {
        public static DataTable ReadExcelFile<T>(string filePath, bool hasHeader, int sheetNumber)
        {
            List<T> rows = new List<T>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTable dataTable = new DataTable();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataset = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = hasHeader
                        }
                    });

                    dataTable = dataset.Tables[sheetNumber - 1];
                }
            }
            return dataTable;
        }

        public static List<T> ReadCSVFile<T>(string filePath, bool hasHeader)
        {
            List<T> rows = new List<T>();
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = hasHeader,
            };
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, csvConfig))
            {
                rows = csv.GetRecords<T>().ToList();
            }
            return rows;
        }
    }
}
