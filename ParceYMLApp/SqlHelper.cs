using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using ParceYmlApp;

namespace ParceYmlApp
{
    public static class SqlHelper
    {

        public static SqlDataReader GetReaderSync(string cnStr, CommandType cType, string cmdText, params SqlParameter[] pList)
        {
            SqlConnection cn = new SqlConnection(cnStr);
            if (cn.State != ConnectionState.Open)
                cn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = cType;
            cmd.CommandTimeout = 6000;
            cmd.CommandText = cmdText;
            if (null != pList)
            {
               // foreach (SqlParameter p in pList)
                //{
                    cmd.Parameters.AddRange(pList);//Add(p);
                //}
            }
            return cmd.ExecuteReader(CommandBehavior.CloseConnection);
        }
        public static DataSet GetDataRequest(string sql, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                // Создаем объект DataAdapter
                SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);
                // Создаем объект Dataset
                DataSet ds = new DataSet();
                // Заполняем Dataset
                adapter.Fill(ds);
                // Отображаем данные
                return ds;
            }
        }
        public static SqlTransaction BeginTransaction(string cnStr)
        {
            SqlConnection cn = new SqlConnection(cnStr);
            cn.Open();
            var trans = cn.BeginTransaction();
            return trans;
        }

        public static SqlTransaction BeginTransaction(SqlConnection cn)
        {
            var trans = cn.BeginTransaction();
            return trans;
        }

        public static async Task<SqlDataReader> GetReader(string cnStr, CommandType cType, string cmdText, params SqlParameter[] pList)
        {
            SqlConnection cn = new SqlConnection(cnStr);
            await cn.OpenAsync();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = cType;
            cmd.CommandTimeout = 6000;
            cmd.CommandText = cmdText;
            if (null != pList)
            {
                foreach (SqlParameter p in pList)
                {
                    cmd.Parameters.Add(p);
                }
            }
            return await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection);
        }

        public static async Task ExecuteCommand(string cnStr, CommandType cType, string cmdText, params SqlParameter[] pList)
        {
            using (SqlConnection cn = new SqlConnection(cnStr))
            {
                await cn.OpenAsync();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandType = cType;
                cmd.CommandTimeout = 6000;
                cmd.CommandText = cmdText;
                if (null != pList)
                {
                    foreach (SqlParameter p in pList)
                    {
                        cmd.Parameters.Add(p);
                    }
                }
                await cmd.ExecuteNonQueryAsync();
            }
        }

        public static int ExecuteCommandSync(string cnStr, CommandType cType, string cmdText, params SqlParameter[] pList)
        {
            using (SqlConnection cn = new SqlConnection(cnStr))
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandType = cType;
                cmd.CommandTimeout = 6000;
                cmd.CommandText = cmdText;
                if (null != pList)
                {
                    foreach (SqlParameter p in pList)
                    {
                        cmd.Parameters.Add(p);
                    }
                }
                return cmd.ExecuteNonQuery();
            }
        }

        public static async Task<XmlDocument> GetXml(string cnStr, CommandType cType, string text, params SqlParameter[] pList)
        {
            XmlDocument result = new XmlDocument();
            using (SqlConnection cn = new SqlConnection(cnStr))
            {
                await cn.OpenAsync();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                cmd.CommandType = cType;
                cmd.CommandText = text;
                if (pList != null)
                {
                    foreach (SqlParameter p in pList)
                    {
                        cmd.Parameters.Add(p);
                    }
                }
                XmlReader r = await cmd.ExecuteXmlReaderAsync();
                result.Load(r);
                cn.Close();
            }
            return result;
        }

        public static async Task InsertBatchAsync(DataTable tbl, string cnStr, int batchSize = 0)
        {
            using (SqlConnection cn = new SqlConnection(cnStr))
            {
                await cn.OpenAsync();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(cn))
                {
                    // Set the batch size and name the destination table 
                    bulkCopy.BatchSize = (batchSize != 0) ? batchSize : tbl.Rows.Count;
                    bulkCopy.DestinationTableName = tbl.TableName;
                    bulkCopy.BulkCopyTimeout = 9000;
                    for (int i = 0; i < tbl.Columns.Count; i++)
                    {
                        bulkCopy.ColumnMappings.Add(tbl.Columns[i].Caption, tbl.Columns[i].Caption);
                    }
                    // Write the data to the destination table 
                    await bulkCopy.WriteToServerAsync(tbl);
                }
            }
        }

        public static DataTable IntArrayToTvpIntArray(this IEnumerable<int> args, string colName = "val")
        {
            var dt = new DataTable();
            dt.Columns.Add(colName, typeof(int));
            if (args != null)
            {
                foreach (var arg in args)
                {
                    dt.Rows.Add(arg);
                }
            }
            return dt;
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
        public static void ExecProc(string SpName)
        {
            try
            {
                using (var connection = new SqlConnection(Program.connectionStr))
                {
                    connection.Open();
                    using (var cmdPost = new SqlCommand())
                    {
                        cmdPost.CommandText = SpName;
                        cmdPost.CommandType = CommandType.StoredProcedure;
                        cmdPost.Connection = connection;
                        Int32 rowsMerge;
                        rowsMerge = cmdPost.ExecuteNonQuery();
                    }
                }

            }
            catch (SqlException se)
            {
                //Console.WriteLine(se);
                MessageBox.Show(se.Message, "Ошибка с сервера!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }

}
