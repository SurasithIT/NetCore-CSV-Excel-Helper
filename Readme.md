# .NET Core CSV and Excel Helper

## Features

- Import CSV and Excel file to DTO (in Excel file can map to DataTable too)
- Export from list of DTO (and from DataTable as Excel file)

## Configuration

- Import
  - with or without header
  - sheet page (only on Excel)
- Export
  - write header (only on CSV file)
  - as Excel file can parse records argument with DataTable or List of DTO type (has 2 overloads function)

## References

- [CsvHelper Document](https://joshclose.github.io/CsvHelper/getting-started/)
- [Convert DataTable to List In C# (c-sharpcorner.com)](https://www.c-sharpcorner.com/UploadFile/ee01e6/different-way-to-convert-datatable-to-list/)
- [c# - How to export DataTable to Excel - Stack Overflow](https://stackoverflow.com/questions/8207869/how-to-export-datatable-to-excel)
- [c# - Exporting the values in List to excel - Stack Overflow](https://stackoverflow.com/questions/2206279/exporting-the-values-in-list-to-excel)
- [ClosedXML](https://github.com/closedxml/closedxml)
- [ExcelDataReader C# (CSharp) Code Examples - HotExamples](https://csharp.hotexamples.com/examples/-/ExcelDataReader/-/php-exceldatareader-class-examples.html)
- [c# - How to change the DataTable Column Name? - Stack Overflow](https://stackoverflow.com/questions/6407239/how-to-change-the-datatable-column-name)
- [Attributes | CsvHelper (joshclose.github.io)](https://joshclose.github.io/CsvHelper/examples/configuration/attributes/)

## Backlog Task

- [x] Import : with Map object by readable specific column name
  - [x] Excel
  - [x] CSV (using attribute)
- [x] Export : with write specific column name as readable
  - [x] Excel
  - [x] CSV (using attribute)
