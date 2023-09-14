using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
                Console.WriteLine($"Executed query: { query}");
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

        public bool RowExist(string NomorAju, string kodeDokumen)
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            string query = $"SELECT * FROM csa.{kodeDokumen}_HEADER WHERE NOMOR_AJU = '{NomorAju}'";

            using(SqlConnection con = new SqlConnection(SambuConnString))
            {
                con.Open();
                using(SqlCommand cmd = new SqlCommand(query, con))
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
        public void RemoveDuplicateColumns(string Path)
        {
            using (var workbook = new XLWorkbook(Path))
            {
                for (int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                {
                    var workSheet = workbook.Worksheet(sh);
                    var columns = new HashSet<string>();
                    var indexColumnToRemove = new HashSet<int>();

                    for (int i = 1; i <= workSheet.LastColumnUsed().ColumnNumber(); i++)
                    {
                        if (!columns.Contains(workSheet.Cell(1, i).Value.ToString()))
                        {
                            columns.Add(workSheet.Cell(1, i).Value.ToString());
                        }

                        else
                        {
                            indexColumnToRemove.Add(i);
                        }


                    }

                    foreach (int colIndex in indexColumnToRemove.OrderByDescending(i => i))
                    {
                        workSheet.Column(colIndex).Delete();
                    }


                }

                workbook.Save();
            }
        }
        public string InsertFromQuery2(string Path)
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            var ExPos = "";
            //string QuerySelect = "";
            string QueryInsert = "";
            string QueryCreate = "";
            string NomorAju = "";

            try
            {
                RemoveDuplicateColumns(Path);
                using (XLWorkbook workbook = new XLWorkbook(Path))
                {
                    string output = "";
                    IXLWorksheet headerSheet = workbook.Worksheet(1);
                    string KodeDokumen = "";
                    KodeDokumen = $"bc{headerSheet.Cell(2, 2).Value}";
                    List<string> headerColumns = new List<string>();
                    #region Assign the value of NomorAju
                    int nomorAjuRow = new();
                    int nomorAjuCol = new();
                    foreach(var cell in headerSheet.Row(1).Cells())
                    {
                        if(cell.Value.ToString() == "NOMOR AJU")
                        {
                            nomorAjuRow = cell.Address.RowNumber;
                            nomorAjuCol = cell.Address.ColumnNumber;
                            break;
                        }
                    }
                    NomorAju = $"{headerSheet.Cell(nomorAjuRow + 1, nomorAjuCol).Value}";
                    #endregion
                    for (int c = 1; c <= headerSheet.LastColumnUsed().ColumnNumber(); c++)
                    {
                        headerColumns.Add(headerSheet.Cell(1, c).Value.ToString());
                    }

                    if (headerColumns.Contains("NOMOR AJU"))
                    {
                        Console.WriteLine($"--Kode Dokumen: {KodeDokumen}");
                        for (int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                        {
                            ExPos = "Sheet (Sheet: " + sh + ") Start";
                            IXLWorksheet sheet = workbook.Worksheet(sh);
                            int SheetRows = sheet.LastRowUsed().RowNumber();
                            int SheetColumns = sheet.LastColumnUsed().ColumnNumber();

                            if (sheet.Name != "VERSI")
                            {
                                #region Get all the column names on each sheet
                                var FirstRow = new List<string>();
                                foreach(var cell in sheet.Row(1).Cells())
                                {
                                    FirstRow.Add(cell.Value.ToString());
                                    
                                }
                                #endregion
                                Console.WriteLine($"Sheet name: {sheet.Name}");
                                Console.WriteLine($"Rows used: {SheetRows}");

                                List<string> Values = new List<string>();
                                

                                //QuerySelect = $"SELECT TOP (1)* FROM csa.{sheet.Name}";
                                QueryInsert = $"INSERT INTO csa.{KodeDokumen}_{sheet.Name.Replace(" ", "_")}\n(\n" + (FirstRow.Contains("NOMOR AJU") ? "" : "[NOMOR_AJU], \n");
                                QueryCreate = $"CREATE TABLE csa.{KodeDokumen}_{sheet.Name.Replace(" ", "_")}\n(\n[ID] INT IDENTITY(1,1),\n" + (FirstRow.Contains("NOMOR AJU") ? "" : "[NOMOR_AJU] VARCHAR(MAX), \n");



                                string InsertValue = "";

                                //Console.WriteLine("--Rows: " + sheet.LastRowUsed().RowNumber());
                                //Console.WriteLine("--Columns: " + sheet.LastColumnUsed().ColumnNumber());

                                string ColumnToCreate = "";
                                HashSet<string> processedColumns = new HashSet<string>();
                                HashSet<string> processedRowsInsert = new HashSet<string>();

                                ExPos += "Sheet(Sheet: " + sh + ") Query Create Table";
                                #region Assign all the column names in the Create table query
                                for (int b = 1; b <= SheetColumns; b++)
                                {
                                    string columnValue = "[" + string.Join("", sheet.Cell(1, b).Value).Replace(" ", "_");
                                    columnValue = columnValue.Replace("/", "_");
                                    columnValue += "]";

                                    if (b == SheetColumns)
                                    {
                                        ColumnToCreate += $"{columnValue} VARCHAR(MAX)\n";
                                    }

                                    else
                                    {
                                        ColumnToCreate += $"{columnValue} VARCHAR(MAX),\n";
                                    }
                                }
                                ColumnToCreate += ")\n";
                                #endregion

                                ExPos = "Sheet(Sheet: " + sh + ") Query Insert Into";
                                #region Assign all the column names in the INSERT INTO query
                                for (int i = 1; i <= SheetColumns; i++)
                                {
                                    string rowValues = "[" + string.Join("", sheet.Cell(1, i).Value).Replace(" ", "_");
                                    rowValues = rowValues.Replace("/", "_");
                                    rowValues += "]";
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

                                if (SheetRows > 1)
                                {
                                    QueryInsert += ")\n";
                                    InsertValue += "\nVALUES\n";
                                }

                                else
                                {
                                    QueryInsert = "";
                                    InsertValue = "";
                                }
                                #endregion

                                #region Assign the values for each column
                                for (int j = 2; j <= SheetRows; j++)
                                {
                                    if ((j != SheetRows))
                                    {
                                        InsertValue += "(" + (FirstRow.Contains("NOMOR AJU") ? "" : NomorAju + ",");
                                        for (int k = 1; k <= SheetColumns; k++)
                                        {
                                            string cellValue = sheet.Cell(j, k).Value.ToString() != null ? sheet.Cell(j, k).Value.ToString() : "";
                                            if (k == SheetColumns)
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
                                        InsertValue += "(" + (FirstRow.Contains("NOMOR AJU") ? "" : NomorAju + ",");
                                        for (int k = 1; k <= SheetColumns; k++)
                                        {
                                            if (k == SheetColumns)
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
                                #endregion

                                bool con = TableExist("csa", $"{KodeDokumen}_{sheet.Name.Replace(" ", "_")}");
                                Console.WriteLine($"Table exists ? : {con}");
                                if (con == true)
                                {
                                    using (SqlConnection conn = new SqlConnection(SambuConnString))
                                    {

                                        conn.Open();
                                        Console.WriteLine($"Checked cell value : {sheet.Cell(2, 1).Value}");
                                        bool exist = RowExist(NomorAju, KodeDokumen);
                                        Console.WriteLine($"Condition met ? : {exist}");
                                        var Query = "";
                                        if (exist == true)
                                        {
                                            var QueryDelete = $"DELETE FROM csa.{KodeDokumen}_{sheet.Name.Replace(" ", "_")} WHERE NOMOR_AJU = '{NomorAju}'";
                                            Query = $"{QueryDelete}\n{(QueryInsert + InsertValue)}";

                                            Console.WriteLine($"Executed Delete and Insert Query : {Query}");
                                            //using (SqlCommand cmd = new SqlCommand(Query, conn))
                                            //{
                                            //    //Debug.WriteLine($"Disini masuk 1");
                                            //    cmd.ExecuteNonQuery();
                                            //    //Debug.WriteLine($"Disini masuk 2");
                                            //}


                                        }

                                        else
                                        {
                                            Query = (QueryInsert + InsertValue);
                                            Console.WriteLine($"Executed query : {Query}");

                                            if (!string.IsNullOrWhiteSpace(Query))
                                            {
                                                //using (SqlCommand cmd = new SqlCommand(Query, conn))
                                                //{
                                                //    //Debug.WriteLine($"Disini masuk 3");
                                                //    cmd.ExecuteNonQuery();
                                                //    //Debug.WriteLine($"Disini masuk 4");
                                                //}
                                            }

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
                                        Console.WriteLine($"Executed Create Query : {(QueryCreate + ColumnToCreate)}");
                                        //using (SqlCommand cmd = new SqlCommand((QueryCreate + ColumnToCreate), conn))
                                        //{
                                        //    cmd.ExecuteNonQuery();
                                        //}

                                        if (!string.IsNullOrWhiteSpace((QueryInsert + InsertValue)))
                                        {
                                            Console.WriteLine($"Executed Insert Query : {(QueryInsert + InsertValue)}");
                                            //using (SqlCommand cmd = new SqlCommand((QueryInsert + InsertValue), conn))
                                            //{
                                            //    Console.WriteLine($"Executed Insert Query : {(QueryInsert + InsertValue)}");
                                            //    cmd.ExecuteNonQuery();
                                            //}
                                        }

                                        else
                                        {
                                            continue;
                                        }



                                        //Console.WriteLine("Table Created and The Data has been inserted");
                                        output = "Table Created and The Data has been inserted";
                                    }
                                }
                            }
                        }
                    }


                    else
                    {
                        output = "HEADER sheet does not have the required column: NOMOR AJU";
                    }



                    return output;
                }
            }

            catch (Exception ex)
            {
                return $"Error: {ex.Message} in {ExPos} /n {QueryCreate} /n {QueryInsert} /n ";
            }



        }
        public string InsertFromQuery3(string Path)
        {
            string SambuConnString = "Data Source=10.0.0.50;Initial Catalog=Sambu_Nintex;User Id=sa; Password=pass@word1";
            var ExPos = "";
            //string QuerySelect = "";
            string QueryInsert = "";
            string QueryCreate = "";
            string NomorAju = "";

            try
            {
                RemoveDuplicateColumns(Path);
                using (XLWorkbook workbook = new XLWorkbook(Path))
                {
                    string output = "";
                    IXLWorksheet headerSheet = workbook.Worksheet(1);
                    string KodeDokumen = "";
                    KodeDokumen = $"bc_testing";
                    List<string> headerColumns = new List<string>();
                    #region Assign the value of NomorAju
                    int nomorAjuRow = new();
                    int nomorAjuCol = new();
                    foreach (var cell in headerSheet.Row(1).Cells())
                    {
                        if (cell.Value.ToString() == "NOMOR AJU")
                        {
                            nomorAjuRow = cell.Address.RowNumber;
                            nomorAjuCol = cell.Address.ColumnNumber;
                            break;
                        }
                    }
                    NomorAju = $"{headerSheet.Cell(nomorAjuRow + 1, nomorAjuCol).Value}";
                    #endregion
                    for (int c = 1; c <= headerSheet.LastColumnUsed().ColumnNumber(); c++)
                    {
                        headerColumns.Add(headerSheet.Cell(1, c).Value.ToString());
                    }

                    if (headerColumns.Contains("NOMOR AJU"))
                    {
                        Console.WriteLine($"--Kode Dokumen: {KodeDokumen}");
                        for (int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                        {
                            ExPos = "Sheet (Sheet: " + sh + ") Start";
                            IXLWorksheet sheet = workbook.Worksheet(sh);
                            int SheetRows = sheet.LastRowUsed().RowNumber();
                            int SheetColumns = sheet.LastColumnUsed().ColumnNumber();

                            if (sheet.Name != "VERSI")
                            {
                                #region Get all the column names on each sheet
                                var FirstRow = new List<string>();
                                foreach (var cell in sheet.Row(1).Cells())
                                {
                                    FirstRow.Add(cell.Value.ToString());

                                }
                                #endregion
                                Console.WriteLine($"Sheet name: {sheet.Name}");
                                Console.WriteLine($"Rows used: {SheetRows}");

                                List<string> Values = new List<string>();


                                //QuerySelect = $"SELECT TOP (1)* FROM csa.{sheet.Name}";
                                QueryInsert = $"INSERT INTO csa.{KodeDokumen}_{sheet.Name.Replace(" ", "_")}\n(\n" + (FirstRow.Contains("NOMOR AJU") ? "" : "[NOMOR_AJU], \n") + (FirstRow.Contains("is_deleted") ? "" : "[is_deleted], \n");
                                QueryCreate = $"CREATE TABLE csa.{KodeDokumen}_{sheet.Name.Replace(" ", "_")}\n(\n[ID] INT IDENTITY(1,1),\n" + (FirstRow.Contains("NOMOR AJU") ? "" : "[NOMOR_AJU] VARCHAR(MAX), \n") + (FirstRow.Contains("id_deleted") ? "" : "[is_deleted] BIT, \n");



                                string InsertValue = "";

                                //Console.WriteLine("--Rows: " + sheet.LastRowUsed().RowNumber());
                                //Console.WriteLine("--Columns: " + sheet.LastColumnUsed().ColumnNumber());

                                string ColumnToCreate = "";
                                HashSet<string> processedColumns = new HashSet<string>();
                                HashSet<string> processedRowsInsert = new HashSet<string>();

                                ExPos += "Sheet(Sheet: " + sh + ") Query Create Table";
                                #region Assign all the column names in the Create table query
                                for (int b = 1; b <= SheetColumns; b++)
                                {
                                    string columnValue = "[" + string.Join("", sheet.Cell(1, b).Value).Replace(" ", "_");
                                    columnValue = columnValue.Replace("/", "_");
                                    columnValue += "]";

                                    if (b == SheetColumns)
                                    {
                                        ColumnToCreate += $"{columnValue} VARCHAR(MAX)\n";
                                    }

                                    else
                                    {
                                        ColumnToCreate += $"{columnValue} VARCHAR(MAX),\n";
                                    }
                                }
                                ColumnToCreate += ")\n";
                                #endregion
                                ExPos = "Sheet(Sheet: " + sh + ") Query Insert Into";
                                #region Assign all the column names in the INSERT INTO query
                                for (int i = 1; i <= SheetColumns; i++)
                                {
                                    string rowValues = "[" + string.Join("", sheet.Cell(1, i).Value).Replace(" ", "_");
                                    rowValues = rowValues.Replace("/", "_");
                                    rowValues += "]";
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

                                if (SheetRows > 1)
                                {
                                    QueryInsert += ")\n";
                                    InsertValue += "\nVALUES\n";
                                }

                                else
                                {
                                    QueryInsert = "";
                                    InsertValue = "";
                                }
                                #endregion
                                #region Assign the values for each column
                                for (int j = 2; j <= SheetRows; j++)
                                {
                                    if ((j != SheetRows))
                                    {
                                        InsertValue += "(" + (FirstRow.Contains("NOMOR AJU") ? "" : NomorAju + ",") + (FirstRow.Contains("is_deleted") ? "" : "0" + ",");
                                        for (int k = 1; k <= SheetColumns; k++)
                                        {
                                            string cellValue = sheet.Cell(j, k).Value.ToString() != null ? sheet.Cell(j, k).Value.ToString() : "";
                                            if (k == SheetColumns)
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
                                        InsertValue += "(" + (FirstRow.Contains("NOMOR AJU") ? "" : NomorAju + ",") + (FirstRow.Contains("is_deleted") ? "" : "0" + ",");
                                        for (int k = 1; k <= SheetColumns; k++)
                                        {
                                            if (k == SheetColumns)
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
                                #endregion
                                #region Check if the table already exists
                                bool con = TableExist("csa", $"{KodeDokumen}_{sheet.Name.Replace(" ", "_")}");
                                Console.WriteLine($"Table exists ? : {con}");
                                #endregion
                                #region If the table already exists, then check the data inside the table based on the value of NOMOR_AJU column
                                if (con == true)
                                {
                                    using (SqlConnection conn = new SqlConnection(SambuConnString))
                                    {
                                        conn.Open();
                                        #region This checks whether the data exists in the table or not
                                        bool exist = RowExist(NomorAju, KodeDokumen);
                                        Console.WriteLine($"Condition met ? : {exist}");
                                        #endregion
                                        var Query = "";
                                        #region If the data exists, then this would delete the data and then re-insert the new data
                                        if (exist == true)
                                        {
                                            var QueryDelete = $"DELETE FROM csa.{KodeDokumen}_{sheet.Name.Replace(" ", "_")} WHERE NOMOR_AJU = '{NomorAju}'";
                                            Query = $"{QueryDelete}\n{(QueryInsert + InsertValue)}";

                                            Console.WriteLine($"Executed Delete and Insert Query : {Query}");
                                            using (SqlCommand cmd = new SqlCommand(Query, conn))
                                            {
                                                //Debug.WriteLine($"Disini masuk 1");
                                                cmd.ExecuteNonQuery();
                                                //Debug.WriteLine($"Disini masuk 2");
                                            }
                                        }
                                        #endregion
                                        #region If the data does not exists, then just insert the data to the table
                                        else
                                        {
                                            Query = (QueryInsert + InsertValue);
                                            Console.WriteLine($"Executed query : {Query}");

                                            if (!string.IsNullOrWhiteSpace(Query))
                                            {
                                                using (SqlCommand cmd = new SqlCommand(Query, conn))
                                                {
                                                    //Debug.WriteLine($"Disini masuk 3");
                                                    cmd.ExecuteNonQuery();
                                                    //Debug.WriteLine($"Disini masuk 4");
                                                }
                                            }
                                        }
                                        #endregion
                                        conn.Close();
                                    }
                                    //Console.WriteLine("Data Inserted Successfully");
                                    output = "Data Inserted Successfully";
                                }
                                #endregion
                                #region If the table does not exists, create a new table and insert data to it
                                else
                                {
                                    using (SqlConnection conn = new SqlConnection((SambuConnString)))
                                    {
                                        conn.Open();
                                        Console.WriteLine($"Executed Create Query : {(QueryCreate + ColumnToCreate)}");
                                        using (SqlCommand cmd = new SqlCommand((QueryCreate + ColumnToCreate), conn))
                                        {
                                            cmd.ExecuteNonQuery();
                                        }

                                        if (!string.IsNullOrWhiteSpace((QueryInsert + InsertValue)))
                                        {
                                            Console.WriteLine($"Executed Insert Query : {(QueryInsert + InsertValue)}");
                                            using (SqlCommand cmd = new SqlCommand((QueryInsert + InsertValue), conn))
                                            {
                                                Console.WriteLine($"Executed Insert Query : {(QueryInsert + InsertValue)}");
                                                cmd.ExecuteNonQuery();
                                            }
                                        }

                                        else
                                        {
                                            continue;
                                        }

                                        //Console.WriteLine("Table Created and The Data has been inserted");
                                        output = "Table Created and The Data has been inserted";
                                    }
                                }
                                #endregion
                            }
                        }
                    }


                    else
                    {
                        output = "HEADER sheet does not have the required column: NOMOR AJU";
                    }
                    return output;
                }
            }

            catch (Exception ex)
            {
                return $"Error: {ex.Message} in {ExPos} /n {QueryCreate} /n {QueryInsert} /n ";
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
