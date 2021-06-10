using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;

namespace NetCore.CSV.Excel.Helper
{
    public class CSVandExcelHelper
    {
        public static DataTable ImportFromExcel(string filePath, bool hasHeader, int sheetNumber = 1)
        {
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

        public static List<T> ImportFromExcel<T>(string filePath, bool hasHeader, int sheetNumber = 1)
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
            rows = ConvertDataTable<T>(dataTable, hasHeader);
            return rows;
        }

        public static List<T> ImportFromCSV<T>(string filePath, bool hasHeader = true)
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

        public static void ExportToCSV<T>(string outputFilePath, List<T> records, bool writeHeader = true)
        {
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = writeHeader
            };
            using (var writer = new StreamWriter(outputFilePath))
            using (var csv = new CsvWriter(writer, csvConfig))
            {
                csv.WriteRecords(records);
            }
        }

        public static void ExportToExcel(string outputFilePath, DataTable dataTable, string sheetName = "Sheet1")
        {
            XLWorkbook workbook = new XLWorkbook();
            workbook.AddWorksheet(sheetName);
            var worksheet = workbook.Worksheet(sheetName);
            worksheet.FirstRow().FirstCell().InsertTable(dataTable);
            workbook.SaveAs(outputFilePath);
        }

        public static void ExportToExcel<T>(string outputFilePath, List<T> records, string sheetName = "Sheet1")
        {
            XLWorkbook workbook = new XLWorkbook();
            workbook.AddWorksheet(sheetName);
            var worksheet = workbook.Worksheet(sheetName);
            worksheet.FirstRow().FirstCell().InsertTable<T>(records);
            workbook.SaveAs(outputFilePath);
        }

        private static List<T> ConvertDataTable<T>(DataTable dataTable, bool hasHeaderRow = true)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dataTable.Rows)
            {
                T item = GetItem<T>(row, hasHeaderRow);
                data.Add(item);
            }
            return data;
        }

        private static T GetItem<T>(DataRow dataRow, bool hasHeaderRow = true)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            if (hasHeaderRow == true)
            {
                //Map by property and column name
                foreach (DataColumn column in dataRow.Table.Columns)
                {
                    foreach (PropertyInfo pro in temp.GetProperties())
                    {
                        if (pro.Name == column.ColumnName)
                        {
                            var data = dataRow[column.ColumnName];
                            var _data = Convert.ChangeType(data, pro.PropertyType);
                            pro.SetValue(obj, _data, null);
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            else
            {
                //Map by property and column index
                for (int i = 0; i < dataRow.Table.Columns.Count; i++)
                {
                    var properties = temp.GetProperties();
                    for (int j = 0; j < properties.Count(); j++)
                    {
                        if (i == j)
                        {
                            var data = dataRow[dataRow.Table.Columns[i].ColumnName];
                            var property = properties[j];
                            var _data = Convert.ChangeType(data, property.PropertyType);
                            property.SetValue(obj, _data, null);
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            return obj;
        }
    }
}
