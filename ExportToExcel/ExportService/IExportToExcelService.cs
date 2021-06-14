namespace ExportToExcel.ExportService
{
    using System.Collections.Generic;
    using System.Data;

    public interface IExportToExcelService
    {
        byte[] CreateExcelDocument<T>(string name, IList<T> list);

        byte[] CreateExcelDocument(string name, DataTable dataTable);

        DataTable CovertDataToDataTable<T>(IEnumerable<T> list);
    }
}