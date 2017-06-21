using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace zinatic_UED.Controllers
{
    public class ExcelController : Controller
    {
        //Variables Excel
        private Excel.Application app = null;
        private Excel.Workbook workbook = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range workSheet_range = null;

        // GET: Excel
        public ActionResult Index()
        {
            return View();
        }

        public void CreateExcel()
        {
            
        }
        public void createDoc()
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            finally
            {
            }
        }

    }
}