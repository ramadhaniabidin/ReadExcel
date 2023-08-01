using ReadExcel;
using System.Data;
using System.Data.SqlClient;

public class Program
{
    public string connString = "Data Source=(localdb)\\local;Initial Catalog=ExcelTest;Integrated Security=True;";
    DatabaseManager db = new DatabaseManager();
    SqlConnection conn = new SqlConnection();


    public static void Main()
    {
        string Path = "C:\\Users\\Bidin\\Downloads\\test.xlsx";

        Console.WriteLine("Hello World");
        ExcelManager excelManager = new ExcelManager();
        DatabaseManager databaseManager = new DatabaseManager();
        DataTable dt = new DataTable();

        dt = excelManager.ExcelRead(Path);
        Console.WriteLine("Column Names:");

        List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName.Replace(" ", "_")).ToList();
        List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();

        Console.WriteLine($"Column count: {Columns.Count}");
        Console.WriteLine($"Row count: {Rows.Count}\n");


        string query = $"INSERT INTO TABLE (\n{string.Join(",\n", Columns)}" + "\n)";
        //Console.WriteLine(query);

        string value = "\n\nVALUES \n";

        //Console.WriteLine("VALUES");
        for (int i = 0; i < Rows.Count; i++)
        {
            //Console.Write("(");
            value += "(";
            for (int j = 0; j < Columns.Count; j++)
            {
                if(j == Columns.Count - 1)
                {
                    value += $"'{dt.Rows[i][j]}'";
                    //Console.Write($"'{dt.Rows[i][j]}'");
                }

                else
                {
                    value += $"'{dt.Rows[i][j]}',";
                    //Console.Write($"'{dt.Rows[i][j]}',");
                }
                
            }

            if(i == Rows.Count - 1)
            {
                value += ")\n\n";
                //Console.WriteLine(")\n");
            }

            else
            {
                value += "),\n\n";
                //Console.WriteLine("),\n");
            }
           
        }

        Console.WriteLine($"\nThe amount of sheet: { excelManager.Test(Path)}");

        InsertFromQuery();
    }


    public void InsertFromSP()
    {
        ExcelManager excelManager = new ExcelManager();
        DataTable dt = new DataTable();
        string Path = "C:\\Users\\Bidin\\Downloads\\test.xlsx";
        dt = excelManager.ExcelRead(Path);

        List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName).ToList();
        List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();

        db.OpenConnection(ref conn);
        for(int i = 0; i < Rows.Count; i++)
        {
            db.cmd.CommandText = "dbo.Insert_Entity_Table";
            db.cmd.CommandType = CommandType.StoredProcedure;
            db.cmd.Parameters.Clear();
            for (int j = 0; j < Columns.Count; j++)
            {

            }
        }





    }

    public static void InsertFromQuery()
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