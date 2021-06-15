namespace ExportToExcel
{
    using System;
    using System.IO;
    using System.Linq;
    using AutoFixture;
    using AutoFixture.Kernel;
    using ExportToExcel.Dto;
    using ExportToExcel.ExportService;

    internal static class Program
    {
        public static void Main(string[] args)
        {
            var fixture = new Fixture();
            // fixture.Customizations.Add(
            //     new TypeRelay(
            //         typeof(FooDto),
            //         typeof(PaymentTypeEnum)));

            var records = fixture
                .CreateMany<FooDto>(count: 100)
                .ToList();

            var exportToExcelService = new ExportToExcelService();
            
            var result = exportToExcelService.WriteToExcel("Excel Export", records);
            
            var excelFile = new FileInfo(@"C:\ExcelExport\exportToExcelDemo.xlsx");
            
            result.SaveAs(excelFile);
            
            Console.WriteLine("Finished.");

        }
    }
}