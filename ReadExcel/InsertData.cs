using ClosedXML.Excel;
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

        public bool TableExist(string schema, string tableName)
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            string query = $"SELECT TOP (1) TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '{schema}' AND TABLE_NAME = '{tableName}'";

            using(SqlConnection conn = new SqlConnection(SambuConnString))
            {
                conn.Open();
                using(SqlCommand cmd = new SqlCommand(query, conn))
                {
                    object res = cmd.ExecuteScalar();
                    if(res != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }


        public void TestSchema()
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            string query = "SELECT TOP (1) TABLE_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'csa'";

            using(SqlConnection conn = new SqlConnection(SambuConnString))
            {
                conn.Open();
                using(SqlCommand cmd = new SqlCommand(query, conn))
                {
                    object res = cmd.ExecuteScalar();
                    if(res != null)
                    {
                        string tableName = res.ToString();
                        Console.WriteLine($"Table Name: {tableName}");
                    }

                    else
                    {
                        Console.WriteLine("No Table Found");
                    }
                }
            }
        }


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
                        db.AddInParameter(db.cmd, paramName, paramValue);
                    }
                    //Console.WriteLine($"Params : \n - " + string.Join("\n - ", db.cmd.Parameters.Cast<SqlParameter>().Select(x => $"{x.ParameterName}: {x.Value}")));
                    db.cmd.ExecuteNonQuery();
                }
                db.trans.Commit();
                db.CloseConnection(ref conn);
                Console.WriteLine("Data inserted successfully.");
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        //public string InsertFromSP1(string Path)
        //{
        //    try
        //    {
        //        db.OpenConnection(ref conn);


        //        db.trans.Commit();
        //        db.CloseConnection(ref conn );
        //    }

        //    catch(Exception ex)
        //    {

        //    }
        //}

        public string InsertFromQuery1(string Path)
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            try
            {
                using (XLWorkbook workbook = new XLWorkbook(Path))
                {
                    string output = "";
                    IXLWorksheet headerSheet = workbook.Worksheet(1);
                    string KodeDokumen = "";
                    KodeDokumen = $"bc{headerSheet.Cell(2, 2).Value}";


                    Console.WriteLine($"--Kode Dokumen: {KodeDokumen}");

                    for (int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                    {
                        IXLWorksheet sheet = workbook.Worksheet(sh);
                        int SheetRows = sheet.LastRowUsed().RowNumber();
                        int SheetColumns = sheet.LastColumnUsed().ColumnNumber();

                        if (SheetRows == 1)
                        {
                            SheetRows += 1;
                        }

                        List<string> Values = new List<string>();

                        string QuerySelect = $"SELECT TOP (1)* FROM csa.{sheet.Name}";
                        string QueryInsert = $"INSERT INTO csa.{KodeDokumen}_{sheet.Name}\n(\n";
                        string QueryCreate = $"CREATE TABLE csa.{KodeDokumen}_{sheet.Name}\n(\n";

                        string InsertValue = "";

                        Console.WriteLine("--Rows: " + sheet.LastRowUsed().RowNumber());
                        Console.WriteLine("--Columns: " + sheet.LastColumnUsed().ColumnNumber());

                        string ColumnToCreate = "";
                        HashSet<string> processedColumns = new HashSet<string>();
                        HashSet<string> processedRowsInsert = new HashSet<string>();

                        for (int b = 1; b <= SheetColumns; b++)
                        {
                            string columnValue = string.Join("", sheet.Cell(1, b).Value).Replace(" ", "_");
                            if (!processedColumns.Contains(columnValue))
                            {
                                processedColumns.Add(columnValue);
                                if (b == SheetColumns)
                                {
                                    ColumnToCreate += $"{columnValue} VARCHAR(MAX)\n";
                                }

                                else
                                {
                                    ColumnToCreate += $"{columnValue} VARCHAR(MAX),\n";
                                }
                            }
                        }

                        ColumnToCreate += ")\n";

                        //Console.WriteLine(QueryInsert);
                        for (int i = 1; i <= SheetColumns; i++)
                        {
                            string rowValues = string.Join("", sheet.Cell(1, i).Value).Replace(" ", "_");
                            if (!processedRowsInsert.Contains(rowValues))
                            {
                                processedRowsInsert.Add(rowValues);
                                if (i == SheetColumns)
                                {
                                    QueryInsert += $"{rowValues}\n";
                                }

                                else
                                {
                                    QueryInsert += $"{rowValues},\n";
                                }
                            }
                        }

                        InsertValue += ")\nVALUES\n";
                        for (int j = 2; j <= SheetRows; j++)
                        {
                            if ((j != SheetRows))
                            {
                                InsertValue += "(";
                                for (int k = 1; k <= processedRowsInsert.Count; k++)
                                {
                                    string cellValue = sheet.Cell(j, k).Value.ToString() != null ? sheet.Cell(j, k).Value.ToString() : "";

                                    if (k == processedRowsInsert.Count)
                                    {
                                        InsertValue += $"'{cellValue}'";
                                    }

                                    else
                                    {
                                        InsertValue += $"'{cellValue}',";
                                    }
                                }
                                InsertValue += "),\n\n";
                            }

                            else
                            {
                                InsertValue += "(";
                                //Console.Write("(");
                                for (int k = 1; k <= processedRowsInsert.Count; k++)
                                {
                                    if (k == processedRowsInsert.Count)
                                    {
                                        InsertValue += $"'{sheet.Cell(j, k).Value}'";
                                    }

                                    else
                                    {
                                        InsertValue += $"'{sheet.Cell(j, k).Value}',";
                                    }
                                }
                                InsertValue += ")\n\n";
                            }
                        }

                        //Console.WriteLine($"\n\n--Select Query: \n{(QuerySelect)}");
                        //Console.WriteLine($"\n\n--Insert Query: \n{(QueryInsert + InsertValue)}");
                        //Console.WriteLine($"\n--Query Create:\n{QueryCreate + ColumnToCreate}");

                        bool con = TableExist("csa", $"{KodeDokumen}_{sheet.Name}");
                        //Console.WriteLine(con);

                        if (con == true)
                        {
                            using (SqlConnection conn = new SqlConnection(SambuConnString))
                            {
                                conn.Open();
                                using (SqlCommand cmd = new SqlCommand((QueryInsert + InsertValue), conn))
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                conn.Close();
                            }
                            //Console.WriteLine("Data Inserted Successfully");
                            output = "Data Inserted Successfully";
                        }

                        else
                        {
                            using (SqlConnection conn = new SqlConnection((SambuConnString)))
                            {
                                conn.Open();
                                using (SqlCommand cmd = new SqlCommand((QueryCreate + ColumnToCreate), conn))
                                {
                                    cmd.ExecuteNonQuery();
                                }

                                using (SqlCommand cmd = new SqlCommand((QueryInsert + InsertValue), conn))
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                output = "Table Created and The Data has been inserted";
                                //Console.WriteLine("Table Created and The Data has been inserted");
                            }
                        }

                    }

                    return output;

                }

                

                //Console.WriteLine("Success");
            }

            catch(Exception ex)
            {
                return $"Error: {ex.Message}";
                //Console.WriteLine("Error: " + ex.Message);
            }
        }

        public string InsertFromQuery2(string Path)
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            try
            {
                using (XLWorkbook workbook = new XLWorkbook(Path))
                {
                    string output = "";
                    IXLWorksheet headerSheet = workbook.Worksheet(1);
                    string KodeDokumen = "";
                    KodeDokumen = $"bc{headerSheet.Cell(2, 2).Value}";


                    Console.WriteLine($"--Kode Dokumen: {KodeDokumen}");

                    for (int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                    {
                        IXLWorksheet sheet = workbook.Worksheet(sh);
                        int SheetRows = sheet.LastRowUsed().RowNumber();
                        int SheetColumns = sheet.LastColumnUsed().ColumnNumber();

                        Console.WriteLine($"Sheet name: {sheet.Name}");
                        Console.WriteLine($"Rows used: {SheetRows}");

                        //if (SheetRows == 1)
                        //{
                        //    SheetRows += 1;
                        //}

                        List<string> Values = new List<string>();

                        string QuerySelect = $"SELECT TOP (1)* FROM csa.{sheet.Name}";
                        string QueryInsert = $"INSERT INTO csa.{KodeDokumen}_{sheet.Name}\n(\n";
                        string QueryCreate = $"CREATE TABLE csa.{KodeDokumen}_{sheet.Name}\n(\n";

                        string InsertValue = "";

                        Console.WriteLine("--Rows: " + sheet.LastRowUsed().RowNumber());
                        Console.WriteLine("--Columns: " + sheet.LastColumnUsed().ColumnNumber());

                        string ColumnToCreate = "";
                        HashSet<string> processedColumns = new HashSet<string>();
                        HashSet<string> processedRowsInsert = new HashSet<string>();

                        for (int b = 1; b <= SheetColumns; b++)
                        {
                            string columnValue = string.Join("", sheet.Cell(1, b).Value).Replace(" ", "_");
                            if (!processedColumns.Contains(columnValue))
                            {
                                processedColumns.Add(columnValue);
                                if (b == SheetColumns)
                                {
                                    ColumnToCreate += $"{columnValue} VARCHAR(MAX)\n";
                                }

                                else
                                {
                                    ColumnToCreate += $"{columnValue} VARCHAR(MAX),\n";
                                }
                            }
                        }

                        ColumnToCreate += ")\n";

                        //Console.WriteLine(QueryInsert);
                        for (int i = 1; i <= SheetColumns; i++)
                        {
                            string rowValues = string.Join("", sheet.Cell(1, i).Value).Replace(" ", "_");
                            if (!processedRowsInsert.Contains(rowValues))
                            {
                                processedRowsInsert.Add(rowValues);
                                if (SheetRows > 1)
                                {
                                    if (i == SheetColumns)
                                    {
                                        QueryInsert += $"{rowValues}\n";
                                    }

                                    else
                                    {
                                        QueryInsert += $"{rowValues},\n";
                                    }
                                }

                                else
                                {
                                    QueryInsert = "";
                                }

                                
                            }

                        }

                        if(SheetRows > 1)
                        {
                            QueryInsert += ")\n";

                            InsertValue += "\nVALUES\n";
                        }

                        else
                        {
                            QueryInsert ="";

                            InsertValue = "";
                        }


                        for (int j = 2; j <= SheetRows; j++)
                        {
                            if ((j != SheetRows))
                            {
                                InsertValue += "(";
                                for (int k = 1; k <= processedRowsInsert.Count; k++)
                                {
                                    string cellValue = sheet.Cell(j, k).Value.ToString() != null ? sheet.Cell(j, k).Value.ToString() : "";

                                    if (k == processedRowsInsert.Count)
                                    {
                                        InsertValue += $"'{cellValue}'";
                                    }

                                    else
                                    {
                                        InsertValue += $"'{cellValue}',";
                                    }
                                }
                                InsertValue += "),\n\n";
                            }

                            else
                            {
                                InsertValue += "(";
                                //Console.Write("(");
                                for (int k = 1; k <= processedRowsInsert.Count; k++)
                                {
                                    if (k == processedRowsInsert.Count)
                                    {
                                        InsertValue += $"'{sheet.Cell(j, k).Value}'";
                                    }

                                    else
                                    {
                                        InsertValue += $"'{sheet.Cell(j, k).Value}',";
                                    }
                                }
                                InsertValue += ")\n\n";
                            }

                        }

                        //Console.WriteLine($"\n\n--Select Query: \n{(QuerySelect)}");
                        Console.WriteLine($"\n\n--Insert Query: \n{(QueryInsert + InsertValue)}");
                        //Console.WriteLine($"\n--Query Create:\n{QueryCreate + ColumnToCreate}");

                        //bool con = TableExist("csa", $"{KodeDokumen}_{sheet.Name}");
                        //Console.WriteLine(con);

                        //if (con == true)
                        //{
                        //    using (SqlConnection conn = new SqlConnection(SambuConnString))
                        //    {
                        //        conn.Open();
                        //        using (SqlCommand cmd = new SqlCommand((QueryInsert + InsertValue), conn))
                        //        {
                        //            cmd.ExecuteNonQuery();
                        //        }
                        //        conn.Close();
                        //    }
                        //    //Console.WriteLine("Data Inserted Successfully");
                        //    output = "Data Inserted Successfully";
                        //}

                        //else
                        //{
                        //    using (SqlConnection conn = new SqlConnection((SambuConnString)))
                        //    {
                        //        conn.Open();
                        //        using (SqlCommand cmd = new SqlCommand((QueryCreate + ColumnToCreate), conn))
                        //        {
                        //            cmd.ExecuteNonQuery();
                        //        }

                        //        using (SqlCommand cmd = new SqlCommand((QueryInsert + InsertValue), conn))
                        //        {
                        //            cmd.ExecuteNonQuery();
                        //        }
                        //        output = "Table Created and The Data has been inserted";
                        //        //Console.WriteLine("Table Created and The Data has been inserted");
                        //    }
                        //}

                    }

                    return output;

                }



                //Console.WriteLine("Success");
            }

            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
                //Console.WriteLine("Error: " + ex.Message);
            }
        }

        //public void InsertFromQuery3(string Path)
        //{
        //    var InsertQuery = "INSERT INTO Table (Col1, Col2, Col3)";
        //    var InsertValue = "\nVALUES";

        //    string[,] value =
        //    {
        //        { "Val11", "Val12", "Val13" },
        //        { "Val21", "Val22", "Val23" },
        //        { "Val31", "Val32", "Val33" }
        //    };


        //}

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
