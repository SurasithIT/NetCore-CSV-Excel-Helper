using System;
using CsvHelper.Configuration.Attributes;

namespace NetCore.CSV.Excel.Helper.Models
{
    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
