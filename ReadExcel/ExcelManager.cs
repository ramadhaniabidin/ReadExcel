using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
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
                    //for(int j = 0; j < worksheet.RowCount(); j++)
                    //{
                    //    Console.WriteLine(j);
                    //}
                }


            }
        }

        public void ExcelMultipleSheets1(string Path)
        {
            using (XLWorkbook workbook = new XLWorkbook(Path))
            {
                for(int sh = 1; sh <= workbook.Worksheets.Count; sh++)
                {
                    IXLWorksheet sheet = workbook.Worksheet(sh);

                    List<string> Values = new List<string>();

                    string QuerySelect = $"SELECT TOP (1)* FROM tmp.{sheet.Name}";
                    string QueryInsert = $"INSERT INTO tmp.{sheet.Name}\n(\n";
                    string QueryCreate = $"CREATE TABLE tmp.{sheet.Name}\n(\n";

                    string InsertValue = "";

                    Console.WriteLine("Rows: " + sheet.LastRowUsed().RowNumber());
                    Console.WriteLine("Columns: " + sheet.LastColumnUsed().ColumnNumber());

                    string ColumnToCreate = "";
                    for(int b = 1; b <= sheet.LastColumnUsed().ColumnNumber(); b++)
                    {
                        if (b == sheet.LastColumnUsed().ColumnNumber())
                        {
                            ColumnToCreate += $"{string.Join("", sheet.Cell(1, b).Value).Replace(" ", "_")} VARCHAR(MAX)\n";
                        }

                        else
                        {
                            ColumnToCreate += $"{string.Join("", sheet.Cell(1, b).Value).Replace(" ", "_")} VARCHAR(MAX),\n";
                        }

                        
                    }

                    ColumnToCreate += ")\n";

                    //Console.WriteLine(QueryInsert);
                    for (int i = 1; i <= sheet.LastColumnUsed().ColumnNumber(); i++)
                    {
                        if (i == sheet.LastColumnUsed().ColumnNumber())
                        {
                            QueryInsert += $"{string.Join("", sheet.Cell(1, i).Value).Replace(" ", "_")}\n";
                        }

                        else
                        {
                            QueryInsert += $"{string.Join("", sheet.Cell(1, i).Value).Replace(" ", "_")},\n";
                        }

                        //Console.WriteLine(string.Join("\n,", sheet.Cell(1, i).Value).Replace(" ", "_"));
                    }

                    InsertValue += ")\nVALUES\n";
                    var tempVal = "";
                    //Console.WriteLine(")\nVALUES\n");
                    for (int j = 2; j <= sheet.LastRowUsed().RowNumber(); j++)
                    {
                        if ((j != sheet.LastRowUsed().RowNumber()))
                        {
                            //Console.Write("(\n");
                            tempVal = "(";
                            InsertValue += "(";
                            for (int k = 1; k <= sheet.LastColumnUsed().ColumnNumber(); k++)
                            {
                                if (k == sheet.LastColumnUsed().ColumnNumber())
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}'";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'");
                                    //Values.Add($"'{sheet.Cell(j, k).Value}'");
                                    //Console.Write($"'{sheet.Cell(j, k).Value}'");
                                }

                                else
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}',";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'" + ", ");
                                    //Values.Add($"'{sheet.Cell(j, k).Value}',");
                                    //Console.Write($"'{sheet.Cell(j, k).Value}',");
                                }
                                //Console.WriteLine($"'{sheet.Cell(j, k).Value}',");
                            }
                            tempVal += ")";
                            InsertValue += "),\n\n";
                            //Console.WriteLine("),\n");

                            Values.Add(tempVal);
                        }

                        else
                        {
                            tempVal = "(";
                            InsertValue += "(";
                            //Console.Write("(");
                            for (int k = 1; k <= sheet.LastColumnUsed().ColumnNumber(); k++)
                            {
                                if (k == sheet.LastColumnUsed().ColumnNumber())
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}'";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'");
                                    //Values.Add($"'{sheet.Cell(j, k).Value}'");
                                    //Console.Write($"'{sheet.Cell(j, k).Value}'");
                                }

                                else
                                {
                                    InsertValue += $"'{sheet.Cell(j, k).Value}',";
                                    tempVal += ($"'{sheet.Cell(j, k).Value}'" + ", ");
                                    //Values.Add($"'{sheet.Cell(j, k).Value}',");
                                    //Console.Write($"'{sheet.Cell(j, k).Value}',");
                                }
                            }
                            tempVal += ")";
                            InsertValue += ")\n\n";
                            //Console.WriteLine(")");
                            Values.Add(tempVal);
                        }
                    }

                    //Console.WriteLine($"\n\nInsert Query: \n{(QueryInsert + InsertValue)}");
                    Console.WriteLine($"\nQuery Create: {QueryCreate + ColumnToCreate}");
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
