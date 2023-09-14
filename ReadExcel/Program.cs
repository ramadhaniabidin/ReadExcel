using ClosedXML.Excel;
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
        string Path = "C:\\Users\\ramad\\Downloads\\BC 1.1 BM Voy 49 SN.xlsx";

        Console.WriteLine("Hello World");
        ExcelManager excelManager = new ExcelManager();
        DatabaseManager databaseManager = new DatabaseManager();
        DataTable dt = new DataTable();

        List<string> columns = new List<string> { "NOMOR AJU", "is_deleted", "NOMOR PABEAN" };
        //bool con = columns.Contains("NOMOR AJU");
        //Console.WriteLine($"Columns : {string.Join(", ", columns)}");

        Console.WriteLine($"{insertData.InsertFromQuery3(Path)}");

        //excelManager.RemoveDuplicateColumns(Path);

    }



}