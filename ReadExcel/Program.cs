using ReadExcel;
using System.Data;

public class Program
{
    public static void Main()
    {
        string Path = "C:\\Users\\Bidin\\Downloads\\test.xlsx";

        Console.WriteLine("Hello World");
        ExcelManager excelManager = new ExcelManager();
        DataTable dt = new DataTable();

        dt = excelManager.ExcelRead(Path);
        Console.WriteLine("Column Names:");

        List<string> Columns = dt.Columns.Cast<DataColumn>().Select(col => col.ColumnName).ToList();
        List<string> Rows = dt.Rows.Cast<DataRow>().Select(row => string.Join("\t", row.ItemArray)).ToList();

        Console.WriteLine($"Column count: {Columns.Count}");
        Console.WriteLine($"Row count: {Rows.Count}\n");


        string query = $"INSERT INTO TABLE (\n{string.Join(",\n", Columns)}" + "\n)";
        Console.WriteLine(query);

        string value = "\n\nVALUES \n";

        Console.WriteLine("VALUES");
        for (int i = 0; i < Rows.Count; i++)
        {
            Console.Write("(");
            value += "(";
            for (int j = 0; j < Columns.Count; j++)
            {
                value += $"'{dt.Rows[i][j]},'";

                if(j == Columns.Count - 1)
                {
                    Console.Write($"'{dt.Rows[i][j]}'");
                }

                else
                {
                    Console.Write($"'{dt.Rows[i][j]}',");
                }
                
            }

            if(i == Rows.Count - 1)
            {
                value += ")\n\n";
                Console.WriteLine(")\n");
            }

            else
            {
                value += "),\n\n";
                Console.WriteLine("),\n");
            }
           
        }
        Console.WriteLine($"\nThe amount of sheet: { excelManager.Test(Path)}");
    }
}