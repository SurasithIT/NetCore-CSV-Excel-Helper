using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;

namespace NetCore.CSV.Excel.Helper.Test
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
            DataTable patientsFromXls = CSVandExcelHelper.ReadExcelFile<Person>(excelFilePath, true, 1);

            string excelFilePathNoHeader = Path.Combine(fileRootPath, excelFileNameWitoutHeader);
            DataTable patientsFromXlsNoHeader = CSVandExcelHelper.ReadExcelFile<Person>(excelFilePathNoHeader, false, 1);

            string csvFilePath = Path.Combine(fileRootPath, csvFileName);
            List<Person> patientsFromCsvWithHeader = CSVandExcelHelper.ReadCSVFile<Person>(csvFilePath, true);

            string csvFilePathNoHeader = Path.Combine(fileRootPath, csvFileNameWithoutHeader);
            List<Person> patientsFromCsvNoHeader = CSVandExcelHelper.ReadCSVFile<Person>(csvFilePathNoHeader, false);
        }
    }

    class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
    }
}