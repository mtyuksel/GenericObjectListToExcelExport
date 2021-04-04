using System;
using System.Collections.Generic;

namespace GenericObjectListToExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creating customer list
            var customers = new List<Customer>();
            
            //Adding some customers
            customers.Add(new Customer
            {
                ID = 1,
                Firstname = "Art",
                Lastname = "Behind Code"
            });
            customers.Add(new Customer
            {
                ID = 2,
                Firstname = "Hello",
                Lastname = "World"
            });

            //Export list as excel
            Helpers.ExportToExcel<Customer>(customers, "CustomerList");
        }
    }
}
