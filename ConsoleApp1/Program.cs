using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace ConsoleApp1
{
    public class Program
    {
        static void Main(string[] args)
        {
          DataTable dt =  Import();
          Export(dt);
        }

        private static DataTable Import()
        {
            string path = @"E:\tmp\XXX.xlsx";
            string sheetName = "Sheet1";
            DataTable dt = new DataTable();
            using (OleDbConnection conn = new OleDbConnection())
            {
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path.Replace(@"\\", @"\") + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    //comm.CommandText = "Select * from [" + sheetName + "$] where [Name] = 'Khandokar'";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                    }
                }
                //EnumerableRowCollection<DataRow> datarows = dt.AsEnumerable().Where(r => r.Field<double>("Id") == 1);
               
            }
            return dt;
        }

        private static void Export(DataTable dt)
        {
            try
            {
                string path = @"E:\tmp1\sabbir.xlsx";
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path.Replace(@"\\", @"\") + ";" +
                    "Mode=ReadWrite;Extended Properties='Excel 12.0 Xml;HDR=YES;MaxScanRows=0;IMEX=0'";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        List<String> columnNames = new List<string>();
                        foreach (DataColumn dataColumn in dt.Columns)
                        {
                            columnNames.Add(dataColumn.ColumnName);
                        }

                        String tableName = !String.IsNullOrWhiteSpace(dt.TableName) ? dt.TableName : Guid.NewGuid().ToString();
                        command.CommandText = $"CREATE TABLE [{tableName}] ({String.Join(",", columnNames.Select(c => $"[{c}] VARCHAR").ToArray())});";
                        command.ExecuteNonQuery();


                        foreach (DataRow row in dt.Rows)
                        {
                            List<String> rowValues = new List<string>();
                            foreach (DataColumn column in dt.Columns)
                            {
                                rowValues.Add((row[column] != null && row[column] != DBNull.Value) ? row[column].ToString() : String.Empty);
                            }
                            command.CommandText = $"INSERT INTO [{tableName}]({String.Join(",", columnNames.Select(c => $"[{c}]"))}) VALUES ({String.Join(",", rowValues.Select(r => $"'{r}'").ToArray())});";
                            command.ExecuteNonQuery();
                        }
                    }

                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    
    }
}
