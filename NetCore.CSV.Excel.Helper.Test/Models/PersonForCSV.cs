using System;
using CsvHelper.Configuration.Attributes;

namespace NetCore.CSV.Excel.Helper.Models
{
    public class PersonForCSV
    {
        [Name("ลำดับ")]
        public int Id { get; set; }

        [Name("ชื่อ")]
        public string Name { get; set; }

        [Name("อายุ")]
        public int Age { get; set; }
    }
}
