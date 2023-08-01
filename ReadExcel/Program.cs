using ReadExcel;
using System.Data;

public class Program
{
    public static void Main()
    {
        Console.WriteLine("Hello World");
        ExcelManager excelManager = new ExcelManager();
        DataTable dt = excelManager.ExcelRead("C:\\Users\\ramad\\Downloads\\test.xlsx");

        Console.WriteLine(dt);
    }
}