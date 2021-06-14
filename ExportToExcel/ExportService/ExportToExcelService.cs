namespace ExportToExcel.ExportService
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.IO;
    using OfficeOpenXml;

    public class ExportToExcelService : IExportToExcelService
    {
       public ExcelPackage WriteToExcel<T>(string name, IEnumerable<T> list)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var excelWorksheet = excelPackage.Workbook.Worksheets.Add(name);
                var dataTable = this.CovertDataToDataTable(list);

                excelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, PrintHeaders: true);

                ApplyHeadingFormatting(excelWorksheet, dataTable);
                ApplyDocumentFormatting(excelWorksheet, dataTable);
                AutoFitColumns(excelWorksheet, dataTable.Columns.Count);

                return ByteArrayToObject(excelPackage.GetAsByteArray());
            }
        }       
       
       public byte[] CreateExcelDocument<T>(string name, IList<T> list)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var excelWorksheet = excelPackage.Workbook.Worksheets.Add(name);
                var dataTable = this.CovertDataToDataTable(list);

                excelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, PrintHeaders: true);

                ApplyHeadingFormatting(excelWorksheet, dataTable);
                ApplyDocumentFormatting(excelWorksheet, dataTable);
                AutoFitColumns(excelWorksheet, dataTable.Columns.Count);

                return excelPackage.GetAsByteArray();
            }
        }

        public byte[] CreateExcelDocument(string name, DataTable dataTable)
        {
            using (var excelPackage = new ExcelPackage())
            {
                var excelWorksheet = excelPackage.Workbook.Worksheets.Add(name);

                excelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, PrintHeaders: true);

                ApplyHeadingFormatting(excelWorksheet, dataTable);
                ApplyDocumentFormatting(excelWorksheet, dataTable);
                AutoFitColumns(excelWorksheet, dataTable.Columns.Count);

                return excelPackage.GetAsByteArray();
            }
        }

        private static ExcelPackage ByteArrayToObject(byte[] arrBytes)
        {
            using (var memStream = new MemoryStream(arrBytes))
            {
                var package = new ExcelPackage(memStream);
                return package;
            }
        }
        
        public DataTable CovertDataToDataTable<T>(IEnumerable<T> list)
        {
            var props = TypeDescriptor.GetProperties(typeof(T));
            var table = new DataTable();

            for (var i = 0; i < props.Count; i++)
            {
                table.Columns.Add(props[i].Name, Nullable.GetUnderlyingType(props[i].PropertyType) ?? props[i].PropertyType);
            }

            var values = new object[props.Count];

            foreach (var item in list)
            {
                for (var i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item) ?? DBNull.Value;
                }

                table.Rows.Add(values);
            }

            return table;
        }

        private static void AutoFitColumns(ExcelWorksheet excelWorksheet, int numberOfColumns)
        {
            for (var i = 1; i <= numberOfColumns; i++)
            {
                excelWorksheet.Column(i).AutoFit();
            }
        }

        private static void ApplyHeadingFormatting(ExcelWorksheet excelWorksheet, DataTable dataTable)
        {
            using (var excelRange = excelWorksheet.Cells[FromRow: 1, FromCol: 1, ToRow: 1, dataTable.Columns.Count])
            {
                excelRange.Style.Font.Bold = true;
                // excelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                // excelRange.Style.Fill.BackgroundColor.Theme(Color.Gray);
                // excelRange.Style.Font.Color.SetColor(Color.White);
            }
        }

        private static void ApplyDocumentFormatting(ExcelWorksheet excelWorksheet, DataTable dataTable)
        {
            if (dataTable.Rows.Count == 0)
            {
                return;
            }

            for (var i = 0; i < dataTable.Columns.Count; i++)
            {
                var type = dataTable.Columns[i].DataType;

                if (type == typeof(DateTime))
                {
                    using (var excelRange = excelWorksheet.Cells[FromRow: 2, i + 1, dataTable.Rows.Count + 1, i + 1])
                    {
                        excelRange.Style.Numberformat.Format = "yyyy/MM/dd";
                    }
                }

                if (type != typeof(double)) continue;
                {
                    using (var excelRange = excelWorksheet.Cells[FromRow: 2, i + 1, dataTable.Rows.Count + 1, i + 1])
                    {
                        excelRange.Style.Numberformat.Format = @"[$R-1C09] #,##0.00";
                    }
                }
            }
        }
    }
}