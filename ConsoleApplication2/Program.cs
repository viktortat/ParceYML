using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ConsoleApplication2.Northwind;
using OfficeOpenXml;
using UtilsTbn;

namespace ConsoleApplication2
{
    class Program
    {
        public IEnumerable<Customer> Customers { get; set; }
        public IEnumerable<Northwind.Order> Orders { get; set; }

        static void Main(string[] args)
        {
            string FullName = @"c:\333\ParceYML\soap.xlsx";
            FileInfo existingFile = new FileInfo(FullName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //                Nom Name    NameTbn Country

                ExcelWorksheet sheet = package.Workbook.Worksheets["Бренды"];
                //var dtBrands = UtilsTbn.Utils.GetDTNew(FullName, 3, "Бренды");
                //var Coll = sheet.Cells[sheet.Dimension.Address]
                //var rowCount = UtilsTbn.Utils.GetLastUsedRow(sheet);
                //var сolCount = sheet.Dimension.End.Column;

                //var Coll = sheet.Cells[1, 1, 40, 4].Rows;
                //var query1 = from cell in Coll
                //        select cell
                //    ;

                /*
                var query1 = (from cell in sheet.Cells["d:d"]
                              where cell.Value is double && (double)cell.Value >= 9990 && (double)cell.Value <= 10000 select cell);

                var query2 = (from cell in sheet.Cells[sheet.Dimension.Address]
                              where cell.Style.Font.Bold select cell);

                var query3 = (from cell in sheet.Cells["d:d"]
                              where cell.Value is double &&
                                    (double)cell.Value >= 9500 && (double)cell.Value <= 10000 &&
                                    cell.Offset(0, -1).GetValue<DateTime>().Year == DateTime.Today.Year + 1
                              select cell);
                */

                //DataTable dt = ToDataTable(dtBrands.ToList());
            }

            /*
            IList<Northwind.Customer> data =
            new Northwind.NorthwindEntities(new Uri("http://services.odata.org/northwind/northwind.svc/"))
            .Customers
            .Expand("Orders")
            .Expand("Orders/Order_Details")
            .Expand("Orders/Order_Details/Product")
            .ToList();
            */



        }

        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                dataTable.Columns.Add(prop.Name); //Setting column names as Property names
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    try
                    {
                        values[i] = Props[i].GetValue(item, null); //inserting property values to datatable rows
                    }
                    catch (Exception ex)
                    {

                    }
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
    }
}
