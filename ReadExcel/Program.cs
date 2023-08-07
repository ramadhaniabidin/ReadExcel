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
        Console.WriteLine(insertData.InsertFromQuery1(Path));
        //insertData.TestSchema();

        bool con = insertData.TableExist("csa", "nota_timbang_header");

        Console.WriteLine($"Output: {insertData.InsertFromQuery2(Path)}");
    }



}