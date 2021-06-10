using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;
using Ganss.Excel;
using Microsoft.Extensions.Configuration;

namespace NetCore_CSV_Excel_Helper
{
    class Program
    {
        static void Main(string[] args)
        {
            var env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
            var builder = new ConfigurationBuilder()
                .AddJsonFile($"appsettings.json", true, true)
                .AddJsonFile($"appsettings.{env}.json", true, true)
                .AddEnvironmentVariables();

            var config = builder.Build();

            string fileRootPath = config.GetValue<string>("RootFilePath");
            string excelFileName = "ex_with_header.xlsx";
            string excelFileNameWitoutHeader = "ex_without_header.xlsx";
            string csvFileName = "ex_with_header.csv";
            string csvFileNameWithoutHeader = "ex_without_header.csv";

            string excelFilePath = Path.Combine(fileRootPath, excelFileName);
            DataTable patientsFromXls2 = ReadExcelFile<Person>(excelFilePath, true, 1);

            string excelFilePathNoHeader = Path.Combine(fileRootPath, excelFileNameWitoutHeader);
            DataTable patientsFromXlsNoHeader2 = ReadExcelFile<PersonNoHeader>(excelFilePathNoHeader, false, 1);

            string csvFilePath = Path.Combine(fileRootPath, csvFileName);
            List<Person> patientsFromCsvWithHeader = ReadCSVFile<Person>(csvFilePath, true).ToList();

            string csvFilePathNoHeader = Path.Combine(fileRootPath, csvFileNameWithoutHeader);
            List<PersonNoHeader> patientsFromCsvNoHeader = ReadCSVFile<PersonNoHeader>(csvFilePathNoHeader, false).ToList();
        }

        static DataTable ReadExcelFile<T>(string filePath, bool hasHeader, int sheetPage)
        {
            List<T> rows = new List<T>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTable dt = new DataTable();
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

                    dt = dataset.Tables[sheetPage - 1];
                }
            }
            return dt;
        }

        static IEnumerable<T> ReadCSVFile<T>(string filePath, bool hasHeader)
        {
            IEnumerable<T> rows = new List<T>();
            var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = hasHeader,
            };
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, csvConfig))
            {
                rows = csv.GetRecords<T>().ToList();
            }
            return rows.AsEnumerable<T>();
        }
    }

    class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class PersonNoHeader
    {
        [Column(1)]
        public int Id { get; set; }
        [Column(2)]
        public string Name { get; set; }
        [Column(3)]
        public int Age { get; set; }
    }
}