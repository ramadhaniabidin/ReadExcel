using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class InsertData
    {
        DatabaseManager db = new DatabaseManager();
        SqlConnection conn = new SqlConnection();

        public void InsertFromSP()
        {

            ExcelManager excelManager = new ExcelManager();
            DataTable dt = new DataTable();
            string Path = "C:\\Users\\ramad\\OneDrive\\Documents\\Belajar\\ReadExcel\\test.xlsx";
            dt = excelManager.ExcelRead(Path);

            List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName.Replace(" ", "_")).ToList();
            List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();

            try
            {
                db.OpenConnection(ref conn);
                for (int i = 0; i < Rows.Count; i++)
                {
                    db.cmd.CommandText = "dbo.Insert_Table_ENTITAS";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();
                    //db.AddInParameter(db.cmd, "NOMOR_AJU", dt.Rows[0][0]);
                    for (int j = 0; j < Columns.Count; j++)
                    {


                        string paramName = $"{Columns[j]}";
                        object paramValue = dt.Rows[i][j];
                        //if (((paramName is string valStr) && (string.IsNullOrWhiteSpace(valStr))) || (paramValue == null))
                        //{
                        //    paramValue = DBNull.Value;
                        //}
                        db.AddInParameter(db.cmd, paramName, paramValue);


                        //Console.WriteLine($"{db.cmd.Parameters[j]} = {paramValue}");

                    }
                    Console.WriteLine($"Params : \n - " + string.Join("\n - ", db.cmd.Parameters.Cast<SqlParameter>().Select(x => $"{x.ParameterName}: {x.Value}")));
                    db.cmd.ExecuteNonQuery();
                    //Console.WriteLine(db.cmd.Parameters);
                }
                //db.trans.Commit();
                db.CloseConnection(ref conn);


                Console.WriteLine("Data inserted successfully.");
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }

        public void InsertFromQuery()
        {
            ExcelManager excelManager = new ExcelManager();
            DataTable dt = new DataTable();
            string Path = "C:\\Users\\ramad\\OneDrive\\Documents\\Belajar\\ReadExcel\\test.xlsx";
            dt = excelManager.ExcelRead(Path);

            List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName.Replace(" ", "_")).ToList();
            List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();
            string connString = "Data Source=(localdb)\\local;Initial Catalog=ExcelTest;Integrated Security=True;";

            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";

            string query = $"INSERT INTO tmp.ENTITAS (\n{string.Join(",\n", Columns)}" + "\n)";
            string value = "\n\nVALUES \n";

            for (int i = 0; i < Rows.Count; i++)
            {
                value += "(";
                for (int j = 0; j < Columns.Count; j++)
                {


                    if (j == Columns.Count - 1)
                    {
                        value += $"'{dt.Rows[i][j]}'";
                    }

                    else
                    {
                        value += $"'{dt.Rows[i][j]}',";
                    }

                }

                if (i == Rows.Count - 1)
                {
                    value += ")\n\n";
                }

                else
                {
                    value += "),\n\n";
                }

            }

            Console.WriteLine(query + value);

            try
            {
                using (SqlConnection conn = new SqlConnection(SambuConnString))
                {
                    conn.Open();


                    using (SqlCommand cmd = new SqlCommand((query + value), conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
                Console.WriteLine("Data inserted successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
