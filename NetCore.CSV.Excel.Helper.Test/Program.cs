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
using NetCore.CSV.Excel.Helper.Models;

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
            //DataTable personsFromXls = CSVandExcelHelper.ReadExcelFile(excelFilePath, true, 1);
            List<Person> personsFromXls = CSVandExcelHelper.ImportFromExcel<Person>(excelFilePath, true, 1);

            string excelFilePathNoHeader = Path.Combine(fileRootPath, excelFileNameWitoutHeader);
            //DataTable personsFromXlsNoHeader = CSVandExcelHelper.ReadExcelFile(excelFilePathNoHeader, false, 1);
            List<Person> personsFromXlsNoHeader = CSVandExcelHelper.ImportFromExcel<Person>(excelFilePathNoHeader, false, 1);

            string csvFilePath = Path.Combine(fileRootPath, csvFileName);
            List<Person> personsFromCsvWithHeader = CSVandExcelHelper.ImportFromExcel<Person>(csvFilePath, true);

            string csvFilePathNoHeader = Path.Combine(fileRootPath, csvFileNameWithoutHeader);
            List<Person> personsFromCsvNoHeader = CSVandExcelHelper.ImportFromCSV<Person>(csvFilePathNoHeader, false);

            string outputCsvPath = Path.Combine(fileRootPath, "output_with_header.csv");
            CSVandExcelHelper.ExportToCSV<Person>(outputCsvPath, personsFromCsvWithHeader, true);

            string outputCsvPathWithoutHeader = Path.Combine(fileRootPath, "output_without_header.csv");
            CSVandExcelHelper.ExportToCSV<Person>(outputCsvPathWithoutHeader, personsFromCsvWithHeader, false);

            string outputXlsPath = Path.Combine(fileRootPath, "output_with_header.xlsx");
            CSVandExcelHelper.ExportToExcel<Person>(outputXlsPath, personsFromXlsNoHeader);
        }
    }
}