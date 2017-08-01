using System.Data;
using System.Data.OleDb;

namespace ExcelReadWrite
{
    public class ExcelReadWrite
    {
        public static DataTable GetDataTable(string sql, string connectionString)
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                // Get all Sheets in Excel File
                DataTable dt = new DataTable();
                cmd.CommandText = sql;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                conn.Close();
                return dt;
            }
        }
    }
}