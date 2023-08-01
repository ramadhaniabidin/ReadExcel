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
            string Path = "C:\\Users\\Bidin\\Downloads\\test.xlsx";
            dt = excelManager.ExcelRead(Path);

            List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName.Replace(" ", "_")).ToList();
            List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();

            try
            {

                for (int i = 0; i < Rows.Count; i++)
                {
                    db.OpenConnection(ref conn);
                    db.cmd.CommandText = "dbo.Insert_Table_ENTITAS";
                    db.cmd.CommandType = CommandType.StoredProcedure;
                    db.cmd.Parameters.Clear();
                    //db.AddInParameter(db.cmd, "NOMOR_AJU", dt.Rows[0][0]);
                    for (int j = 0; j < Columns.Count; j++)
                    {


                        string paramName = $"{Columns[j]}";
                        object paramValue = dt.Rows[i][j];
                        db.AddInParameter(db.cmd, paramName, paramValue);


                        Console.WriteLine($"{db.cmd.Parameters[j]} = {paramValue}");

                    }

                    db.cmd.ExecuteNonQuery();
                    db.CloseConnection(ref conn);
                    //Console.WriteLine(db.cmd.Parameters);



                }


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
            string Path = "C:\\Users\\Bidin\\Downloads\\test.xlsx";
            dt = excelManager.ExcelRead(Path);

            List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName.Replace(" ", "_")).ToList();
            List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();
            string connString = "Data Source=(localdb)\\local;Initial Catalog=ExcelTest;Integrated Security=True;";

            string query = $"INSERT INTO dbo.ENTITAS (\n{string.Join(",\n", Columns)}" + "\n)";
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
                using (SqlConnection conn = new SqlConnection(connString))
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
