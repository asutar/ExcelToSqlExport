using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ExcelToSqlExport.Models
{
    public sealed class DataAccessLayer
    {
        string conExcelString = string.Empty;
        string conSqlString = string.Empty;
        public DataTable dt = new DataTable();

        public DataAccessLayer()
        {
            conExcelString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
            conSqlString = ConfigurationManager.ConnectionStrings["Constring"].ConnectionString;
        }
    
        public DataTable GetDataTable()
        {
           
            DataTable dt = new DataTable();

            using (OleDbConnection connExcel = new OleDbConnection(conExcelString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand("Select * from [sheet1$]", connExcel))
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter(cmdExcel))
                    {
                        DataTable EmployeeTable = new DataTable();
                        odaExcel.Fill(dt);
                    }
                }
            }

            return dt;
        }
        public void sqlBulkCopy(DataTable dt)
        {
            using (SqlConnection con = new SqlConnection(conSqlString))
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    //Set the database table name.
                    sqlBulkCopy.DestinationTableName = "dbo.TBL_EMPLOYEE";

                    //[OPTIONAL]: Map the Excel columns with that of the database table
                    sqlBulkCopy.ColumnMappings.Add("NAME", "NAME");
                    sqlBulkCopy.ColumnMappings.Add("ADDRESS", "ADDRESS");

                    con.Open();
                    sqlBulkCopy.WriteToServer(dt);
                    con.Close();
                }
            }
        }
    }
}