using Npgsql;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketingDataProcessing.Utilities
{
    class SqlDataAccess
    {
        private readonly string _ConnectionString = "Host=localhost;Port=5432;Username=postgres;Password=vutuanlinh;Database=MarketingDatabase";
        private NpgsqlConnection _NpgsqlConnection = null;
        public SqlDataAccess()
        {
            _ConnectionString = "Host=localhost;Port=5432;Username=postgres;Password=123456a@;Database=MarketingDatabase";
        }
        /// <summary>
        /// Execute a select query to a database.
        /// </summary>
        /// <param name="query">The SQL SELECT statement.</param>
        /// <param name="parameters">A list of parameters, uses the query want to query.</param>
        /// <returns>A data reader object.</returns>

        public DataTable ExecuteSelectQuery(string query, SqlParameter[] parameters)
        {

            using (_NpgsqlConnection = new NpgsqlConnection(_ConnectionString))// connect to the database.
            {
                _NpgsqlConnection.Open();
                using (NpgsqlCommand command = new NpgsqlCommand())  // create a new command to prepare executing a query request.
                {
                    command.Connection = _NpgsqlConnection;
                    command.CommandText = query;

                    // add parameters for query.
                    if (parameters != null && parameters.Length != 0)
                    {
                        foreach (var param in parameters)
                        {
                            command.Parameters.AddWithValue(param.ParameterName, param.Value);
                        }
                    }
                    NpgsqlDataReader dataReader = command.ExecuteReader(); // execute a query to the database.
                    DataTable dataTable = new DataTable();
                    dataTable.Fill(dataReader);
                    return dataTable;
                }
            }
        }
        public DataTable ExecuteSelectMultiTables(string query, SqlParameter[] parameters)
        {

            using (_NpgsqlConnection = new NpgsqlConnection(_ConnectionString))// connect to the database.
            {
                _NpgsqlConnection.Open();
                using (NpgsqlCommand command = new NpgsqlCommand())  // create a new command to prepare executing a query request.
                {
                    command.Connection = _NpgsqlConnection;
                    command.CommandText = query;

                    // add parameters for query.
                    if (parameters != null && parameters.Length != 0)
                    {
                        foreach (var param in parameters)
                        {
                            if (param.Value == DBNull.Value)
                            {
                                command.Parameters.AddWithValue(param.ParameterName, DBNull.Value);
                            }
                            else
                            {
                                command.Parameters.AddWithValue(param.ParameterName, param.Value);
                            }

                        }
                    }
                    NpgsqlDataReader dataReader = command.ExecuteReader(); // execute a query to the database.
                    DataTable dataTable = new DataTable();
                    dataTable.Fill(dataReader);
                    return dataTable;
                }
            }
        }
        public Object ExecuteSelectQuery(string query, params string[] conditions)
        {
            return null;
        }
        /// <summary>
        /// Execute an update or an insert query.
        /// </summary>
        /// <param name="query">The SQL UPDATE(INSERT) statement.</param>
        /// <param name="parameters">>A list of parameters, uses the query want to query.</param>
        /// <returns>The number of rows was effected.</returns>
        public int ExecuteInsertOrUpdateQuery(string query, SqlParameter[] parameters)
        {
            using (_NpgsqlConnection = new NpgsqlConnection(_ConnectionString)) // connect to the database.
            {
                _NpgsqlConnection.Open();
                using (NpgsqlCommand command = new NpgsqlCommand()) // create a new command to prepare executing a query request.
                {
                    command.Connection = _NpgsqlConnection;
                    command.CommandText = query;

                    // add parameters for query.
                    if (parameters != null || parameters.Length != 0)
                    {
                        foreach (var param in parameters)
                        {
                            command.Parameters.AddWithValue(param.ParameterName, param.Value); // execute a query to the database.
                        }
                    }
                    int effectedRows = command.ExecuteNonQuery();
                    return effectedRows;
                }
            }
        }
        /// <summary>
        /// Execute an delete query.
        /// </summary>
        /// <param name="query">The SQL DELETE statement.</param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public bool ExecuteDeleteQuery(string query, SqlParameter[] parameters)
        {
            using (_NpgsqlConnection = new NpgsqlConnection(_ConnectionString)) // connect to the database.
            {
                _NpgsqlConnection.Open();
                using (NpgsqlCommand command = new NpgsqlCommand()) // create a new command to prepare executing a query request.
                {
                    command.Connection = _NpgsqlConnection;
                    command.CommandText = query;

                    // add parameters for query.
                    if (parameters != null || parameters.Length != 0)
                    {
                        foreach (var param in parameters)
                        {
                            command.Parameters.AddWithValue(param.ParameterName, param.Value);
                        }
                    }
                    int effectedRows = command.ExecuteNonQuery(); // execute a query to the database.
                    return effectedRows == 0 ? false : true;
                }
            }
        }
    }
}
