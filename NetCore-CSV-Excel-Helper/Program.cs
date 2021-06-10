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

namespace NetCore_CSV_Excel_Helper
{
    class Program
    {
        static void Main(string[] args)
        {

            string excelFilePath = "telemed_patientview.xls";
            List<Patient> patientsFromXls = ReadExcelFile<Patient>(excelFilePath, true).ToList();

            string excelFilePathNoHeader = "telemed_patientview_no_header.xls";
            List<PatientNoHeader> patientsFromXlsNoHeader = ReadExcelFile<PatientNoHeader>(excelFilePathNoHeader, false).ToList();

            string csvFilePath = "telemed_patientview.csv";
            List<Patient> patientsFromCsvWithHeader = ReadCSVFile<Patient>(csvFilePath, true).ToList();

            string csvFilePathNoHeader = "telemed_patientview_no_header.csv";
            List<PatientNoHeader> patientsFromCsvNoHeader = ReadCSVFile<PatientNoHeader>(csvFilePathNoHeader, false).ToList();

            DataTable patientsFromXls2 = ReadExcelFile<Patient>(excelFilePath, true, 1);
            DataTable patientsFromXlsNoHeader2 = ReadExcelFile<Patient>(excelFilePathNoHeader, false, 1);
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

                    //for (int i = 0; i < dt.Rows.Count; i++)
                    //{
                    //    Console.WriteLine($"Row {i}");
                    //    for (int j = 0; j < dt.Columns.Count; j++)
                    //    {
                    //        Console.Write($"{dt.Rows[i][j]}\t");
                    //    }
                    //    Console.Write("\n");
                    //}


                    //do
                    //{
                    //    //T row = new T();
                    //    while (reader.Read()) //Each ROW
                    //    {
                    //        var row = Activator.CreateInstance<T>();
                    //        PropertyInfo[] properties = row.GetType().GetProperties();
                    //        for (int column = 0; column < reader.FieldCount; column++)
                    //        {
                    //            //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
                    //            Console.Write(reader.GetValue(column) + "\t");//Get Value returns object
                    //            if (modelStructure.ContainsKey(column))
                    //            {
                    //                var _row = row.GetType().GetProperty(modelStructure[column]);
                    //                var val = reader.GetValue(column);
                    //                var rowType = _row.PropertyType;
                    //                var convertedVal = Convert.ChangeType(val, rowType);
                    //                row.GetType().GetProperty(modelStructure[column]).SetValue(row, convertedVal);
                    //            }
                    //        }
                    //        Console.Write("\n");
                    //    }
                    //} while (reader.NextResult()); //Move to NEXT SHEET
                }
            }

            return dt;
        }

        static IEnumerable<T> ReadExcelFile<T>(string filePath, bool hasHeader)
        {
            IEnumerable<T> rows = new List<T>();
            if (hasHeader == false)
            {
                rows = new ExcelMapper(filePath) { HeaderRow = false }.Fetch<T>();
            }
            rows = new ExcelMapper(filePath).Fetch<T>();
            return rows.AsEnumerable<T>();
        }

        static IEnumerable<T> ReadCSVFile<T>(string filePath, bool hasHeader)
        {
            IEnumerable<T> rows = new List<T>();
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = hasHeader,
            };
            using (var stream = new StreamReader(filePath))
            {
                using (var csv = new CsvReader(stream, config))
                {
                    csv.Read();
                    if (hasHeader == true)
                    {
                        csv.ReadHeader();
                    }
                    rows = csv.GetRecords<T>().ToList();
                }
            }
            return rows.AsEnumerable<T>();
        }
    }

    class Patient
    {
        public string HCODE { get; set; }
        public string HN { get; set; }
        public string DOB { get; set; }
        public string MOBILENO { get; set; }
        public string ADDRESSNOTE { get; set; }
        public int? GENDER { get; set; }
        public string FULLNAME { get; set; }
        public string CID { get; set; }
    }

    class PatientNoHeader
    {
        //[Column(1)]
        public int? No { get; set; }
        [Column(2)]
        //[Column(Letter = "B")]
        public string HCODE { get; set; }
        [Column(3)]
        //[Column(Letter = "C")]
        public string HN { get; set; }
        [Column(4)]
        public string DOB { get; set; }
        [Column(5)]
        public string MOBILENO { get; set; }
        [Column(6)]
        public string ADDRESSNOTE { get; set; }
        [Column(7)]
        public int? GENDER { get; set; }
        [Column(8)]
        public string FULLNAME { get; set; }
        [Column(9)]
        public string CID { get; set; }
    }
}