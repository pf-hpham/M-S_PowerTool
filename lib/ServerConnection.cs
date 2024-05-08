using System;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Windows;

namespace MnS.lib
{
    public static class ServerConnection
    {
        public static SqlConnection OpenConnection(string server_link)
        {
            try
            {
                SqlConnection connection = new SqlConnection(server_link);
                connection.Open();
                if (connection.State == ConnectionState.Open)
                {
                    return connection;
                }
                else
                {
                    MessageBox.Show("Cannot connect to server.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        public static void CloseConnection(SqlConnection connection)
        {
            try
            {
                if (connection != null && connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static OdbcConnection OpenODBCConnection(string server_link)
        {
            try
            {
                OdbcConnection connection = new OdbcConnection(server_link);
                connection.Open();
                if (connection.State == ConnectionState.Open)
                {
                    return connection;
                }
                else
                {
                    MessageBox.Show("Cannot connect to server.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
    }
}