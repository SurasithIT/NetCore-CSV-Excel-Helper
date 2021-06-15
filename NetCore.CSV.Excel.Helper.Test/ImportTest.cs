using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NetCore.CSV.Excel.Helper.Models;
using NetCore.CSV.Excel.Helper.Test.Utilities;
using Xunit;

namespace NetCore.CSV.Excel.Helper.Test
{
    public class ImportTest
    {
        [Fact]
        public void ImportFromCSVWithHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.csvFileName);
            List<Person> obj = CSVandExcelHelper.ImportFromCSV<Person>(filePath, true);
            Assert.IsType<List<Person>>(obj);
            List<Person> expect = Common.people;
            Assert.True(checkPeople(expect, obj));
        }

        [Fact]
        public void ImportFromCSVWithoutHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.csvFileNameWithoutHeader);
            List<Person> obj = CSVandExcelHelper.ImportFromCSV<Person>(filePath, false);
            Assert.IsType<List<Person>>(obj);
            List<Person> expect = Common.people;
            Assert.True(checkPeople(expect, obj));

        }

        [Fact]
        public void ImportFromCSVWithCustomHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.csvFileNameCustomHeader);
            List<PersonForCSV> obj = CSVandExcelHelper.ImportFromCSV<PersonForCSV>(filePath, true);
            Assert.IsType<List<PersonForCSV>>(obj);
            List<PersonForCSV> expect = Common.peopleForCSV;
            Assert.True(checkPeople(expect, obj));
        }

        [Fact]
        public void ImportFromExcelWithHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.excelFileName);
            List<Person> obj = CSVandExcelHelper.ImportFromExcel<Person>(filePath, true);
            Assert.IsType<List<Person>>(obj);
            List<Person> expect = Common.people;
            Assert.True(checkPeople(expect, obj));
        }

        [Fact]
        public void ImportFromExcelVWithoutHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.excelFileNameWitoutHeader);
            List<Person> obj = CSVandExcelHelper.ImportFromExcel<Person>(filePath, false);
            Assert.IsType<List<Person>>(obj);
            List<Person> expect = Common.people;
            Assert.True(checkPeople(expect, obj));
        }

        [Fact]
        public void ImportFromExcelWithCustomHeader()
        {
            string filePath = Path.Combine(Common.fileRootPath, Common.excelFileNameCustomHeader);
            List<Person> obj = CSVandExcelHelper.ImportFromExcel<Person>(filePath, true, 1, Common.columnMappers);
            Assert.IsType<List<Person>>(obj);
            List<Person> expect = Common.people;
            Assert.True(checkPeople(expect, obj));
        }

        private bool checkPeople(List<Person> expect, List<Person> actaul)
        {
            try
            {
                if(expect.Count != actaul.Count)
                {
                    return false;
                }
                bool check = true;
                expect = expect.OrderBy(o => o.Id).ToList();
                actaul = actaul.OrderBy(o => o.Id).ToList();

                for (int i = 0;i < expect.Count; i++)
                {
                    bool checkId = expect[i].Id == actaul[i].Id;
                    bool checkName = expect[i].Name == actaul[i].Name;
                    bool checkAge = expect[i].Age == actaul[i].Age;
                    if(checkId == false || checkName == false || checkAge == false)
                    { 
                        check = false;
                        break;
                    }
                }
                return check;
            }catch(Exception ex)
            {
                throw ex;
            }
        }

        private bool checkPeople(List<PersonForCSV> expect, List<PersonForCSV> actaul)
        {
            try
            {
                if (expect.Count != actaul.Count)
                {
                    return false;
                }
                bool check = true;
                expect = expect.OrderBy(o => o.Id).ToList();
                actaul = actaul.OrderBy(o => o.Id).ToList();

                for (int i = 0; i < expect.Count; i++)
                {
                    bool checkId = expect[i].Id == actaul[i].Id;
                    bool checkName = expect[i].Name == actaul[i].Name;
                    bool checkAge = expect[i].Age == actaul[i].Age;
                    if (checkId == false || checkName == false || checkAge == false)
                    {
                        check = false;
                        break;
                    }
                }
                return check;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
