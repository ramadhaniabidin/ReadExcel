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

        //XLWorkbook UpdatedWorkbook = excelManager.RemoveDuplicateColumns1(Path);
        //excelManager.TestRemove(Path);

        //using (XLWorkbook OriginalWorkbook = new XLWorkbook(Path))
        //{

        //    for (int sh = 1; sh <= OriginalWorkbook.Worksheets.Count; sh++)
        //    {
        //        var OWBSheet = OriginalWorkbook.Worksheet(sh);

        //        Console.WriteLine("Original workbook columns: ");
        //        for (int i = 1; i <= OWBSheet.LastColumnUsed().ColumnNumber(); i++)
        //        {
        //            Console.WriteLine(OWBSheet.Cell(1, i).Value);
        //        }

        //        Console.WriteLine();
        //    }
        //}
        //Console.WriteLine();

        //for(int sh = 1; sh <= UpdatedWorkbook.Worksheets.Count; sh++)
        //{
        //    var UpdatedSHeet = UpdatedWorkbook.Worksheet(sh);
        //    Console.WriteLine("Updated workbook columns:");
        //    for(int i = 1; i <= UpdatedSHeet.LastColumnUsed().ColumnNumber(); i++)
        //    {
        //        Console.WriteLine(UpdatedSHeet.Cell(1, i).Value);
        //    }

        //    Console.WriteLine() ;
        //}

        Console.WriteLine($"{insertData.InsertFromQuery2(Path)}");

        //excelManager.RemoveDuplicateColumns(Path);

    }



}