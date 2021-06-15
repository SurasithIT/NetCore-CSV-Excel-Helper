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

namespace NetCore.CSV.Excel.Helper.Program
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
            string csvFileName = "ex_with_header.csv";
            string csvFileNameWithoutHeader = "ex_without_header.csv";
            string csvFileNameCustomHeader = "ex_custom_header.csv";

            string excelFileName = "ex_with_header.xlsx";
            string excelFileNameWitoutHeader = "ex_without_header.xlsx";
            string excelFileNameCustomHeader = "ex_custom_header.xlsx";

            List<Person> people = new List<Person>()
            {
                new Person(){Id = 1, Name = "Jonathan", Age = 19},
                new Person(){Id = 1, Name = "Lionel", Age = 23},
                new Person(){Id = 1, Name = "Antonio", Age = 28},
            };

            List<PersonForCSV> peopleForCSV = new List<PersonForCSV>()
            {
                new PersonForCSV(){Id = 1, Name = "Jonathan", Age = 19},
                new PersonForCSV(){Id = 1, Name = "Lionel", Age = 23},
                new PersonForCSV(){Id = 1, Name = "Antonio", Age = 28},
            };

            List<ColumnMapperModel> columnMappers = new List<ColumnMapperModel>()
            {
                new ColumnMapperModel(){PropertyName = "Id", ColumnName = "ลำดับ"},
                new ColumnMapperModel(){PropertyName = "Name", ColumnName = "ชื่อ"},
                new ColumnMapperModel(){PropertyName = "Age", ColumnName = "อายุ"}
            };

            #region Export
            #region Export to CSV
            string csvFilePath = Path.Combine(fileRootPath, csvFileName);
            CSVandExcelHelper.ExportToCSV<Person>(csvFilePath, people, true);

            string csvNoHeaderFilePath = Path.Combine(fileRootPath, csvFileNameWithoutHeader);
            CSVandExcelHelper.ExportToCSV<Person>(csvNoHeaderFilePath, people, false);

            string csvCustomHeaderFilePath = Path.Combine(fileRootPath, csvFileNameCustomHeader);
            CSVandExcelHelper.ExportToCSV<PersonForCSV>(csvCustomHeaderFilePath, peopleForCSV, true);
            #endregion

            #region Export to Excel
            string excelFilePath = Path.Combine(fileRootPath, excelFileName);
            CSVandExcelHelper.ExportToExcel<Person>(excelFilePath, people);

            string excelNoHeaderFilePath = Path.Combine(fileRootPath, excelFileNameWitoutHeader);
            CSVandExcelHelper.ExportToExcel<Person>(excelNoHeaderFilePath, people, false, columnMappers);

            string excelCustomHeaderFilePath = Path.Combine(fileRootPath, excelFileNameCustomHeader);
            CSVandExcelHelper.ExportToExcel<Person>(excelCustomHeaderFilePath, people, true, columnMappers);
            #endregion
            #endregion

            #region Import
            #region Import from CSV
            List<Person> personsFromCsv = CSVandExcelHelper.ImportFromCSV<Person>(csvFilePath, true);

            List<Person> personsFromCsvNoHeader = CSVandExcelHelper.ImportFromCSV<Person>(csvNoHeaderFilePath, false);

            List<PersonForCSV> personsFromCsvCustomHeader = CSVandExcelHelper.ImportFromCSV<PersonForCSV>(csvCustomHeaderFilePath, true);
            #endregion

            #region Import from Excel
            //DataTable personsFromXls = CSVandExcelHelper.ImportFromExcel(excelFilePath, true, 1);
            List<Person> personsFromXls = CSVandExcelHelper.ImportFromExcel<Person>(excelFilePath, true);

            //DataTable personsFromXlsNoHeader = CSVandExcelHelper.ImportFromExcel(excelFilePathNoHeader, false, 1);
            List<Person> personsFromXlsNoHeader = CSVandExcelHelper.ImportFromExcel<Person>(excelNoHeaderFilePath, false);

            //DataTable personsFromXls1 = CSVandExcelHelper.ImportFromExcel(excelCustomHeaderFilePath, true, 1);
            List<Person> personsFromXlsCustomHeader = CSVandExcelHelper.ImportFromExcel<Person>(excelCustomHeaderFilePath, true, 1, columnMappers);
            #endregion
            #endregion
        }
    }
}