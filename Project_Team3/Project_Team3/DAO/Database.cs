using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace Project_Team4.DAO
{
    class Database
    {

        public static DataTable GetBySQL(string sql, string dbName)
        {
            try {
                string sqlConnection = "Data Source = localhost; Initial Catalog = '" + dbName + "';" +
                    " Integrated Security = SSPI";
                SqlConnection con = new SqlConnection(sqlConnection);
                SqlCommand cmd = new SqlCommand(sql, con);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds.Tables[0];
            }
            catch (Exception) {

            }
            return null;
        }

        public static int ExecuteSQL(string sql, params SqlParameter[] sqlParameters)
        {
            try {
                string sqlConnection = "Data Source = localhost; Initial Catalog = DBI202_PE_Result;" +
                    " Integrated Security = SSPI";
                SqlConnection con = new SqlConnection(sqlConnection);
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddRange(sqlParameters);
                cmd.Connection.Open();
                int result = cmd.ExecuteNonQuery();
                cmd.Connection.Close();
                return result;
            }
            catch (Exception) {

            }
            return 0;
        }

        public static string CreateDB()
        {
            string sql = "CREATE DATABASE DBI202_PE_Result";
            string ConnectionString = "Data Source = localhost; Initial Catalog = master;" +
                    " Integrated Security = SSPI";
            SqlCommand cmd = null;
            SqlConnection con = null;
            con = new SqlConnection(ConnectionString);
            // Open the connecttion
            if (con.State != ConnectionState.Open) {
                con.Open();
            }
            cmd = new SqlCommand(sql, con);
            try {
                cmd.ExecuteNonQuery();
                return "success";
            }
            catch (SqlException e) {
                return e.Message;
            }
        }

        public static string CreateTable(string tableName)
        {
            string sqlConnection = "Data Source = localhost; Initial Catalog = DBI202_PE_Result;" +
                    " Integrated Security = SSPI";
            SqlCommand cmd = null;
            SqlConnection con = null;
            con = new SqlConnection(sqlConnection);
            // Open the connecttion
            if (con.State != ConnectionState.Open) {
                con.Open();
            }
            string sql = "CREATE TABLE [dbo].[" + tableName + "]( [studentID] [VARCHAR](10) NOT NULL" +
                " PRIMARY KEY, [studentName] [NVARCHAR](50) NULL, [examPaperCode] [VARCHAR](5)" +
                " NULL, [mark] [FLOAT] NULL, [markingDetail] [NVARCHAR](1000) NULL, )";
            cmd = new SqlCommand(sql, con);
            try {
                cmd.ExecuteNonQuery();
                return "success";
            }
            catch (SqlException e) {
                return e.Message;
            }
        }
        public static DataTable GetNameDatabase()
        {
            string sql = "SELECT name FROM sys.databases";
            return GetBySQL(sql, "master");
        }

        public static DataTable GetAllTable()
        {
            string sql = "SELECT TABLE_CATALOG, TABLE_NAME FROM DBI202_PE_Result.INFORMATION_SCHEMA.TABLES";
            return GetBySQL(sql, "master");
        }
    }
}
