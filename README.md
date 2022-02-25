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
          </br>
            using (ExcelEngine engine = new ExcelEngine())
            {
              </br>
                IApplication application = engine.Excel;
                  </br>
                application.DefaultVersion = ExcelVersion.Xlsx;
                  </br>

                IWorkbook workbook = application.Workbooks.Create(1);
                  </br>
                IWorksheet sheet = workbook.Worksheets[0];
                  </br>
                DataSet dataSet = new DataSet();
                  </br>
                dataSet.ReadXml(Path.GetFullPath("book.xml"));
                  </br>
                DataTable dataTable = new DataTable();
                  </br>
                dataTable = dataSet.Tables[0];
                  </br>
                sheet.ImportDataTable(dataTable, true, 1, 1, true);
                  </br>
                IListObject table = sheet.ListObjects.Create("book", sheet.UsedRange);
                  </br>
                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleLight4;
                  </br>
                sheet.UsedRange.AutofitColumns();
                  </br>
                Stream excelstream = File.Create(Path.GetFullPath("book.xlsx"));
                  </br>
                workbook.SaveAs(excelstream);
                  </br>
                excelstream.Dispose();
                  </br>
                Console.WriteLine("successfully converted to Excel");
                  </br>
            }
              </br>
        }
          </br>
    }
      </br>
}
