namespace ExportToExcel.Dto
{
    using System;

    public class FooDto
    {
        public PaymentTypeEnum PaymentType { get; set; }
        
        public string Entity { get; set; }
        
        public string SupplierName { get; set; }
        
        public string ExpenseType { get; set; }

        public string Department { get; set; }

        public string BudgetItem { get; set; }
        
        public string ProjectCode { get; set; }

        public string UserApproved { get; set; }
        
        public decimal TotalInvoiceAmount { get; set; }
        
        public decimal InvoiceItemAmount { get; set; }
        
        public string InvoiceItemDescription { get; set; }
        
        public DateTime Date { get; set; }
    }
}