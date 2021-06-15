using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Extensions.Configuration;
using NetCore.CSV.Excel.Helper.Models;
using NetCore.CSV.Excel.Helper.Test.Utilities;
using Xunit;

namespace NetCore.CSV.Excel.Helper.Test
{
    public class ExportTest
    {
        [Fact]
        public void ExportToCSVWithHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.csvFileName);
            CSVandExcelHelper.ExportToCSV<Person>(filePath, Common.people, true);
            Assert.True(File.Exists(filePath));
        }

        [Fact]
        public void ExportToCSVWithoutHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.csvFileNameWithoutHeader);
            CSVandExcelHelper.ExportToCSV<Person>(filePath, Common.people, false);
            Assert.True(File.Exists(filePath));
        }

        [Fact]
        public void ExportToCSVWithCustomHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.csvFileNameCustomHeader);
            CSVandExcelHelper.ExportToCSV<PersonForCSV>(filePath, Common.peopleForCSV, true);
            Assert.True(File.Exists(filePath));
        }

        [Fact]
        public void ExportToExcelWithHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.excelFileName);
            CSVandExcelHelper.ExportToExcel<Person>(filePath, Common.people, true);
            Assert.True(File.Exists(filePath));
        }

        [Fact]
        public void ExportToExcelWithoutHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.excelFileNameWitoutHeader);
            CSVandExcelHelper.ExportToExcel<Person>(filePath, Common.people, false);
            Assert.True(File.Exists(filePath));
        }

        [Fact]
        public void ExportToExcelWithCustomHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.excelFileNameCustomHeader);
            CSVandExcelHelper.ExportToExcel<Person>(filePath, Common.people, true, Common.columnMappers);
            Assert.True(File.Exists(filePath));
        }
    }
}
