using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ExcelManager
    {

        public List<string> GetSheetNames(string Path)
        {
            List<string> SheetNames = new List<string>();
            using(XLWorkbook workbook = new XLWorkbook(Path))
            {
                for(int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    IXLWorksheet worksheet = workbook.Worksheet(i + 1);
                    SheetNames.Add(worksheet.Name);
                }
            }

            return SheetNames;
        }

        public int Test(string Path)
        {
            int sheets;
            
            using(XLWorkbook workbook = new XLWorkbook(Path))
            {
                sheets = workbook.Worksheets.Count;
            }

            return sheets;
        }

        public DataTable ExcelMultipleSheets(string Path)
        {
            using(XLWorkbook workbook = new XLWorkbook())
            {
                DataTable dt = new DataTable();
                for(int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    IXLWorksheet worksheet = workbook.Worksheet(i + 1);

                    bool firstRow = true;
                    int idxRow = 0;
                    foreach(IXLRow row in worksheet.Rows())
                    {
                        if(firstRow)
                        {
                            foreach(IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }

                        else
                        {
                            dt.Rows.Add();
                            int j = 0;
                            foreach(IXLCell cell in row.Cells())
                            {
                                dt.Rows[dt.Rows.Count - 1][j] = cell.Value.ToString();
                            }
                            idxRow++;
                        }
                    }
                }

                return dt;
            }
        }

        public void ExcelMultipleSheets2(string Path)
        {
            
            using (XLWorkbook workbook = new XLWorkbook(Path))
            {
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    IXLWorksheet worksheet = workbook.Worksheet(i + 1);
                    string QuerySelect = $"SELECT TOP (1)* FROM tmp.{worksheet.Name}";
                    string QueryInsert = $"INSERT INTO tmp.{worksheet.Name}";
                    string QueryCreate = $"CREATE TABLE tmp.{worksheet.Name}";
                    Console.WriteLine(QueryInsert);
                    foreach(IXLCell cell in worksheet.Cells())
                    {
                        Console.WriteLine(cell.Value.ToString());
                    }
                    Console.WriteLine();
                }


            }
        }

        public void ExcelMultipleSheets1(string Path)
        {
            using (XLWorkbook workbook = new XLWorkbook(Path))
            {
                IXLWorksheet headerSheet = workbook.Worksheet(1);
                string KodeDokumen = "";
                KodeDokumen = $"bc{headerSheet.Cell(2,2).Value}";
                
                
                Console.WriteLine($"Kode Dokumen: {KodeDokumen}");

                for (int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                {
                    IXLWorksheet sheet = workbook.Worksheet(sh);
                    int SheetRows = sheet.LastRowUsed().RowNumber();
                    int SheetColumns = sheet.LastColumnUsed().ColumnNumber();

                    if(SheetRows == 1)
                    {
                        SheetRows += 1;
                    }

                    List<string> Values = new List<string>();

                    string QuerySelect = $"SELECT TOP (1)* FROM tmp.{sheet.Name}";
                    string QueryInsert = $"INSERT INTO tmp.{sheet.Name}\n(\n";
                    string QueryCreate = $"CREATE TABLE tmp.{KodeDokumen}_{sheet.Name}\n(\n";

                    string InsertValue = "";



                    string ColumnToCreate = "";
                    HashSet<string> processedColumns = new HashSet<string>();
                    HashSet<string> processedRowsInsert = new HashSet<string>();

                    Console.WriteLine("Rows: " + SheetRows);
                    Console.WriteLine("Columns: " + SheetColumns);

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
                    var tempVal = "";
                    //Console.WriteLine(")\nVALUES\n");
                    for (int j = 2; j <= SheetRows; j++)
                    {
                        if (j != SheetRows)
                        {

                            tempVal = "(";
                            InsertValue += "(";
                            for (int k = 1; k <= processedRowsInsert.Count; k++)
                            {

                                if (k == processedRowsInsert.Count)
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}'";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'");
                                }

                                else
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}',";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'" + ", ");
                                }
                            }
                            tempVal += ")";
                            InsertValue += "),\n\n";

                            Values.Add(tempVal);
                        }

                        else
                        {
                            tempVal = "(";
                            InsertValue += "(";
                            for (int k = 1; k <= processedRowsInsert.Count; k++)
                            {
                                if (k == processedRowsInsert.Count)
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}'";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'");
                                }

                                else
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}',";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'" + ", ");
                                }
                            }
                            tempVal += ")";
                            InsertValue += ")\n\n";
                            Values.Add(tempVal);
                        }
                    }

                    Console.WriteLine($"\n\nInsert Query: \n{(QueryInsert + InsertValue)}");
                    //Console.WriteLine($"\nQuery Create: {QueryCreate + ColumnToCreate}");
                }



            }
        }

        public DataTable ExcelRead(string Path)
        {
            //Save the uploaded Excel file.
            //string filePath = Server.MapPath("~/Files/") + Path.GetFileName(FileUpload1.PostedFile.FileName);
            //FileUpload1.SaveAs(filePath);

            //Open the Excel file using ClosedXML.
            using(XLWorkbook workbook = new XLWorkbook(Path))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet worksheet = workbook.Worksheet(2);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                int idxRow = 0;
                foreach(IXLRow row in worksheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if(firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                            //Console.WriteLine(dt);
                        }
                        firstRow = false;
                        
                    }

                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            //Console.WriteLine(dt);
                            i++;
                        }
                        idxRow++;
                        //Console.WriteLine(dt);
                    }

                    
                }
                return dt;
            }
        }


        public DataTable ExcelRead1(string Path)
        {
            //Save the uploaded Excel file.
            //string filePath = Server.MapPath("~/Files/") + Path.GetFileName(FileUpload1.PostedFile.FileName);
            //FileUpload1.SaveAs(filePath);

            //Open the Excel file using ClosedXML.
            using (XLWorkbook workbook = new XLWorkbook(Path))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet worksheet = workbook.Worksheet(2);
                //workbook.Worksheets.Count;

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                int idxRow = 0;
                foreach (IXLRow row in worksheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                            //Console.WriteLine(dt);
                        }
                        firstRow = false;

                    }

                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            //Console.WriteLine(dt);
                            i++;
                        }
                        idxRow++;
                        //Console.WriteLine(dt);
                    }


                }
                return dt;
            }
        }
    }
}
