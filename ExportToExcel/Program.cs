namespace ExportToExcel
{
    using System;
    using System.IO;
    using System.Linq;
    using AutoFixture;
    using ExportToExcel.Dto;
    using ExportToExcel.ExportService;

    internal static class Program
    {
        public static void Main(string[] args)
        {
            var fixture = new Fixture();
            
            var records = fixture
                .CreateMany<FooDto>()
                .ToList();

            var exportToExcelService = new ExportToExcelService();
            
            var result = exportToExcelService.WriteToExcel("Excel Export", records);
            
            var excelFile = new FileInfo(@"C:\ExcelExport\test.xlsx");
            
            result.SaveAs(excelFile);
            
            Console.WriteLine("Finished.");

        }
    }
}