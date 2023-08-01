using ClosedXML.Excel;
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
        public DataTable ExcelRead(string Path)
        {
            //Save the uploaded Excel file.
            //string filePath = Server.MapPath("~/Files/") + Path.GetFileName(FileUpload1.PostedFile.FileName);
            //FileUpload1.SaveAs(filePath);

            //Open the Excel file using ClosedXML.
            using(XLWorkbook workbook = new XLWorkbook(Path))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet worksheet = workbook.Worksheet(1);

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
                            i++;
                        }
                        idxRow++;
                    }
                }
                return dt;
            }
        }
    }
}
