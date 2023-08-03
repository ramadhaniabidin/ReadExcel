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
        InsertData insertData = new InsertData();
        string Path = "C:\\Users\\ramad\\OneDrive\\Documents\\Belajar\\ReadExcel\\bc41.xlsx";

        Console.WriteLine("Hello World");
        ExcelManager excelManager = new ExcelManager();
        DatabaseManager databaseManager = new DatabaseManager();
        DataTable dt = new DataTable();

        dt = excelManager.ExcelRead(Path);
        //excelManager.ExcelMultipleSheets1(Path);
        insertData.InsertFromQuery1(Path);
        //insertData.TestSchema();

        bool con = insertData.TableExist("csa", "nota_timbang_header");

        Console.WriteLine(con);
        //Console.WriteLine("Column Names:");

        List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName.Replace(" ", "_")).ToList();
        List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();

        //Console.WriteLine($"Column count: {Columns.Count}");
        //Console.WriteLine($"Row count: {Rows.Count}\n");


        string query = $"INSERT INTO TABLE (\n{string.Join(",\n", Columns)}" + "\n)";
        //Console.WriteLine(query);

        string value = "\n\nVALUES \n";

        //for (int i = 0; i < Rows.Count; i++)
        //{
        //    Console.WriteLine("Parameters");
        //    for (int j = 0; j < Columns.Count; j++)
        //    {
        //        Console.WriteLine($"{Columns[j]}: {dt.Rows[i][j]}");
        //    }
        //    Console.WriteLine();
        //}


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

        //Console.WriteLine(query + value);
        //Console.WriteLine($"\nThe amount of sheet: { dt}");
        //insertData.InsertFromQuery();
        //insertData.InsertFromSP();
        //List<string> SheetNames = excelManager.GetSheetNames(Path);
        //Console.WriteLine(string.Join("\n", SheetNames));
    }



}