using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using LinqToExcel;

namespace ConsoleApplication1
{
    class Program
    {
        //add linqtoexcel  https://github.com/paulyoder/LinqToExcel
        //+https://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
        //https://www.microsoft.com/en-us/download/details.aspx?id=13255

        static void Main(string[] args)
        {
//            №	Название НазваниеTBN Страна


         var eFilePath = @"c:\333\ParceYML\soap.xlsx";
            var excel = new ExcelQueryFactory(eFilePath);
            var i = 1;
            var oldCompanies = from c in excel.Worksheet<Brand>("Бренды") //worksheet name = 'US Companies'
                               //where c.LaunchDate < new DateTime(1900, 1, 1)
                               select new
                               {
                                   c.Nom,
                                   c.Name,
                                   c.NameTbn,
                                   c.country
                               };
            DataTable dt = ToDataTable(oldCompanies.ToList());
        }


        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                dataTable.Columns.Add(prop.Name);//Setting column names as Property names
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);//inserting property values to datatable rows
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

    }




}
