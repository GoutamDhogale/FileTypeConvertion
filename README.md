# FileTypeConvertion XML to Excel

using System;
</br>
using System.Collections.Generic;
</br>
using System.Data;
</br>
using System.IO;
</br>
using System.Linq;
</br>
using System.Text;
</br>
using System.Threading.Tasks;
</br>
using Syncfusion.XlsIO;
</br>
namespace xmltoxl
</br>
{
</br>
    class Program
    </br>
    {
    </br>
        static void Main(string[] args)
        {
            using (ExcelEngine engine = new ExcelEngine())
            {
                IApplication application = engine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];
                DataSet dataSet = new DataSet();
                dataSet.ReadXml(Path.GetFullPath("book.xml"));
                DataTable dataTable = new DataTable();
                dataTable = dataSet.Tables[0];
                sheet.ImportDataTable(dataTable, true, 1, 1, true);
                IListObject table = sheet.ListObjects.Create("book", sheet.UsedRange);
                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleLight4;
                sheet.UsedRange.AutofitColumns();
                Stream excelstream = File.Create(Path.GetFullPath("book.xlsx"));
                workbook.SaveAs(excelstream);
                excelstream.Dispose();
                Console.WriteLine("successfully converted to Excel");
            }
        }
    }
}
