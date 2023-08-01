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

            }
        }
    }
}
