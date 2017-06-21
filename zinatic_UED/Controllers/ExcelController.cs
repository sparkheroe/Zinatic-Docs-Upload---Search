using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace zinatic_UED.Controllers
{
    public class ExcelController : Controller
    {
        Excel.Application excel;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        Excel.Range cellRange;

        // GET: Excel
        public ActionResult Index()
        {
            CreateExcel();
            return View();
        }
        public System.Data.DataTable ExportToExcel()
        {
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Sex", typeof(string));
            table.Columns.Add("Subject1", typeof(int));
            table.Columns.Add("Subject2", typeof(int));
            table.Columns.Add("Subject3", typeof(int));
            table.Columns.Add("Subject4", typeof(int));
            table.Columns.Add("Subject5", typeof(int));
            table.Columns.Add("Subject6", typeof(int));
            table.Rows.Add(1, "Amar", "M", 78, 59, 72, 95, 83, 77);
            table.Rows.Add(2, "Mohit", "M", 76, 65, 85, 87, 72, 90);
            table.Rows.Add(3, "Garima", "F", 77, 73, 83, 64, 86, 63);
            table.Rows.Add(4, "jyoti", "F", 55, 77, 85, 69, 70, 86);
            table.Rows.Add(5, "Avinash", "M", 87, 73, 69, 75, 67, 81);
            table.Rows.Add(6, "Devesh", "M", 92, 87, 78, 73, 75, 72);
            return table;
        }
        public void CreateExcel()
        {
            excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            workbook = excel.Workbooks.Add(Type.Missing);

            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = "StudentRepoertCard";

            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 8]].Merge();

            int rowcount = 2;

            foreach (DataRow datarow in ExportToExcel().Rows)
            {
                rowcount += 1;
                for (int i = 1; i <= ExportToExcel().Columns.Count; i++)
                {

                    if (rowcount == 3)
                    {
                        worksheet.Cells[2, i] = ExportToExcel().Columns[i - 1].ColumnName;
                        worksheet.Cells.Font.Color = System.Drawing.Color.Black;

                    }

                    worksheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                    if (rowcount > 3)
                    {
                        if (i == ExportToExcel().Columns.Count)
                        {
                            if (rowcount % 2 == 0)
                            {
                                cellRange = worksheet.Range[worksheet.Cells[rowcount, 1], worksheet.Cells[rowcount, ExportToExcel().Columns.Count]];
                            }

                        }
                    }

                }

            }

            cellRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowcount, ExportToExcel().Columns.Count]];
            cellRange.EntireColumn.AutoFit();
            Excel.Borders border = cellRange.Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            cellRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, ExportToExcel().Columns.Count]];

            workbook.SaveAs("C:\\Users\\Emilio\\Documents\\visual studio 2017\\Projects\\zinatic_UED\\zinatic_UED\\files\\archivo45.xls"); ;
            workbook.Close();
            excel.Quit();
        }
    }
}