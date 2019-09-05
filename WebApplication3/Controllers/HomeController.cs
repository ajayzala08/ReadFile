using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using WebApplication3.Models;

namespace WebApplication3.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpGet]
        public ActionResult UploadWordDocument()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ReadTable(HttpPostedFileBase uploadfile)
        {

            uploadfile.SaveAs(Server.MapPath("\\Resume\\") + uploadfile.FileName);
            string filepath = Server.MapPath("\\Resume\\" + uploadfile.FileName.ToString());

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Documents docs = app.Documents;
            Document doc = docs.Open(filepath, ReadOnly: true);
            Table t = doc.Tables[1];
            Range r = t.Range;
            Cells cells = r.Cells;

            tbl_model tmodel = new tbl_model();
            tmodel.name = cells[3].Range.Text.ToString();
            tmodel.cnt = cells[5].Range.Text.ToString();
            tmodel.qualification = cells[7].Range.Text;
            tmodel.email= cells[9].Range.Text;

            for (int i = 1; i <= cells.Count; i++)
            {
                Cell cell = cells[i];
                Range r2 = cell.Range;
                String txt = r2.Text;
                Marshal.ReleaseComObject(cell);
                Marshal.ReleaseComObject(r2);
            }

            //Rows rows = t.Rows;
            //Columns cols = t.Columns;
            // Cannot access individual rows in this collection because the table has vertically merged cells.
            //for (int i = 0; i < rows.Count; i++) {
            //  for (int j = 0; j < cols.Count; j++) {
            //      Cell cell = rows[i].Cells[j];
            //      Range r = cell.Range;
            //  }
            //}

            doc.Close(false);
            app.Quit(false);
            //Marshal.ReleaseComObject(cols);
            //Marshal.ReleaseComObject(rows);
            Marshal.ReleaseComObject(cells);
            Marshal.ReleaseComObject(r);
            Marshal.ReleaseComObject(t);
            Marshal.ReleaseComObject(doc);
            Marshal.ReleaseComObject(docs);
            Marshal.ReleaseComObject(app);
            return View(tmodel);

        }
    }
}