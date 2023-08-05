using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class DatabaseManager
    {
        public string connString = "Data Source=(localdb)\\local;Initial Catalog=ExcelTest;Integrated Security=True;";
        public SqlCommand cmd = new SqlCommand();
        public SqlDataReader dReader;
        public SqlTransaction trans;


        public void OpenConnection(ref SqlConnection connection, bool isTrans = true)
        {
            if(connection == null)
            {
                connection = new SqlConnection(connString);
                connection.Open();
                cmd = connection.CreateCommand();
                cmd.CommandTimeout = 0;
                if (isTrans)
                {
                    trans = connection.BeginTransaction();
                    cmd.Transaction = trans;
                }
            }

            else
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection = new SqlConnection(connString);
                    connection.Open();
                    cmd = connection.CreateCommand();
                    cmd.CommandTimeout = 0;
                    if (isTrans)
                    {
                        trans = connection.BeginTransaction();
                        cmd.Transaction = trans;
                    }
                }
            }
        }


        public void CloseConnection(ref SqlConnection connection, bool isTrans = false)
        {
            if (connection?.State == ConnectionState.Open)
            {
                if (isTrans)
                {
                    trans.Commit();
                }
                connection?.Close();
            }
            connection?.Dispose();
            cmd?.Dispose();
        }

        public void AddInParameter(SqlCommand command, string name, object value)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.ParameterName = name;
            parameter.Value = value;
            parameter.Direction = ParameterDirection.Input;

            command.Parameters.Add(parameter);
        }

        public void AddInParameter(SqlCommand command, string name, SqlDbType type)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.ParameterName = name;
            parameter.SqlDbType = type;
            parameter.Direction = ParameterDirection.Output;

            command.Parameters.Add(parameter);
        }

        public void CloseDataReader(SqlDataReader dataReader)
        {
            if (dataReader == null)
            {
                return;
            }
            dataReader.Close();
            dataReader.Dispose();
        }

        public DataTable GetDataByProcedure(string SPName, string ParamName, string ParamValue)
        {
            DataTable dt = new DataTable();
            SqlDataReader dr = null;
            cmd.CommandText = SPName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue(ParamName, ParamValue);
            dr = cmd.ExecuteReader();
            dt.Load(dr);
            CloseDataReader(dr);
            return dt;
        }

        public bool isRecordExists(string SPName, string ParamName, string ParamValue)
        {
            SqlDataReader dr = null;
            cmd.CommandText = SPName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Clear();
            cmd.Parameters.AddWithValue(ParamName, ParamValue);
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                CloseDataReader(dr);
                return true;
            }
            else
            {
                CloseDataReader(dr);
                return false;
            }
        }

        public bool isRecordExists(string query)
        {
            SqlDataReader dr = null;
            cmd.CommandText = query;
            cmd.CommandType = CommandType.StoredProcedure;
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                CloseDataReader(dr);
                return true;
            }
            else
            {
                CloseDataReader(dr);
                return false;
            }
        }


        public string AutoCounter(string fieldName, string TableName, string fieldCriteria, string valueCriteria, int LengthOfString, string fieldNameConverted = "")
        {
            SqlDataReader dataReader;
            string str = "";
            string str2 = "1";

            cmd.CommandText = $"SELECT TOP 1 {fieldName} FROM {TableName} WHERE {fieldCriteria} LIKE '%{valueCriteria}%'";
            if (fieldNameConverted != "")
            {
                cmd.CommandText += $" ORDER BY {fieldNameConverted} DESC";
            }
            else
            {
                cmd.CommandText += $" ORDER BY {fieldName} DESC";
            }
            cmd.CommandType = CommandType.Text;
            dReader = cmd.ExecuteReader();

            if (dReader.HasRows)
            {
                dReader.Read();
                str2 = dReader[fieldName].ToString();

                int StartFrom = valueCriteria.Length;
                int EndUntil = 0;
                EndUntil = str2.Length - StartFrom;

                str2 = (int.Parse(str2.Substring(StartFrom, EndUntil)) + 1).ToString();
            }

            int i = 0;
            while (i < LengthOfString - str2.Length)
            {
                str += "0";
                i += 1;
            }
            str = valueCriteria + str + str2;
            CloseDataReader(dReader);
            return str;
        }

        public string GetValueFromQuery(string Query, string FieldName)
        {
            string result = "";
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = Query;
            dReader = cmd.ExecuteReader();
            while (dReader.Read())
            {
                result = dReader[FieldName].ToString();
            }
            CloseDataReader(dReader);
            return result;
        }

        public DataTable GetValueFromSP(string SPName, string ParamName, string ParamValue)
        {
            DataTable dt = new DataTable();
            cmd.CommandText = SPName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Clear();
            AddInParameter(cmd, ParamName, ParamValue);
            dReader = cmd.ExecuteReader();
            dt.Load(dReader);
            CloseDataReader(dReader);
            return dt;
        }

    }
}
