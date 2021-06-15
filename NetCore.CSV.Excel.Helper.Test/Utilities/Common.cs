using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using NetCore.CSV.Excel.Helper.Models;

namespace NetCore.CSV.Excel.Helper.Test.Utilities
{
    public class Common
    {
        public static IConfiguration config = new ConfigurationBuilder()
                                            .SetBasePath(AppContext.BaseDirectory)
                                            .AddJsonFile("appsettings.json", false, true)
                                            .Build();

        public static string fileRootPath = config.GetValue<string>("RootFilePath");
        public static string csvFileName = "ex_with_header.csv";
        public static string csvFileNameWithoutHeader = "ex_without_header.csv";
        public static string csvFileNameCustomHeader = "ex_custom_header.csv";

        public static string excelFileName = "ex_with_header.xlsx";
        public static string excelFileNameWitoutHeader = "ex_without_header.xlsx";
        public static string excelFileNameCustomHeader = "ex_custom_header.xlsx";

        public static List<Person> people = new List<Person>()
        {
            new Person(){Id = 1, Name = "Jonathan", Age = 19},
            new Person(){Id = 2, Name = "Lionel", Age = 23},
            new Person(){Id = 3, Name = "Antonio", Age = 28},
        };

        public static List<PersonForCSV> peopleForCSV = new List<PersonForCSV>()
        {
            new PersonForCSV(){Id = 1, Name = "Jonathan", Age = 19},
            new PersonForCSV(){Id = 2, Name = "Lionel", Age = 23},
            new PersonForCSV(){Id = 3, Name = "Antonio", Age = 28},
        };

        public static List<ColumnMapperModel> columnMappers = new List<ColumnMapperModel>()
        {
            new ColumnMapperModel(){PropertyName = "Id", ColumnName = "ลำดับ"},
            new ColumnMapperModel(){PropertyName = "Name", ColumnName = "ชื่อ"},
            new ColumnMapperModel(){PropertyName = "Age", ColumnName = "อายุ"}
        };
    }
}
