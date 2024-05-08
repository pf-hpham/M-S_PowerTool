using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace MnS.lib
{
    public static class SQLDataTool
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="query"></param>
        /// <param name="parameters"></param>
        /// <param name="server_link"></param>
        /// <returns></returns>
        public static DataTable QueryUserData(string query, List<SqlParameter> parameters, string server_link)
        {
            SqlConnection connection = ServerConnection.OpenConnection(server_link);
            DataTable dt = new DataTable();
            try
            {
                using (SqlCommand sqlCmd = new SqlCommand(query, connection))
                {
                    sqlCmd.Parameters.AddRange(parameters.ToArray());
                    using (SqlDataAdapter da = new SqlDataAdapter(sqlCmd))
                    {
                        da.Fill(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while querying data: " + ex.Message);
            }
            finally
            {
                ServerConnection.CloseConnection(connection);
            }
            return dt;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="query"></param>
        /// <param name="parameters"></param>
        /// <param name="server_link"></param>
        public static void ExecuteNonQuery(string query, List<SqlParameter> parameters, string server_link)
        {
            SqlConnection connection = ServerConnection.OpenConnection(server_link);
            try
            {
                using (SqlCommand sqlCmd = new SqlCommand(query, connection))
                {
                    sqlCmd.Parameters.AddRange(parameters.ToArray());
                    sqlCmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error while executing non-query: " + ex.Message);
            }
            finally
            {
                ServerConnection.CloseConnection(connection);
            }
        }
    }
}
