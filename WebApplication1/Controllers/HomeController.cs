using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;
using Microsoft.Office;
using System.IO;
using System.Text.RegularExpressions;
using WebApplication1.Models;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace WebApplication1.Controllers
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
        
        public ActionResult ReadWordFilePage()
        {
            return View();
        }

        public ActionResult UploadandReadWordFile(HttpPostedFileBase uploadfile)
        {
            object documentFormat = 8;
            string randomName = DateTime.Now.Ticks.ToString();
            object htmlFilePath = Server.MapPath("~/Temp/") + randomName + ".htm";
            string directoryPath = Server.MapPath("~/Temp/") + randomName + "_files";
            object fileSavePath = Server.MapPath("~/Temp/") + Path.GetFileName(uploadfile.FileName);

            //If Directory not present, create it.
            if (!Directory.Exists(Server.MapPath("~/Temp/")))
            {
                Directory.CreateDirectory(Server.MapPath("~/Temp/"));
            }

            //Upload the word document and save to Temp folder.
            uploadfile.SaveAs(fileSavePath.ToString());

            //Open the word document in background.
            _Application applicationclass = new Application();
            applicationclass.Documents.Open(ref fileSavePath);
            applicationclass.Visible = false;
            Document document = applicationclass.ActiveDocument;

            //Save the word document as HTML file.
            document.SaveAs(ref htmlFilePath, ref documentFormat);

            //Close the word document.
            document.Close();

            //Read the saved Html File.
            string wordHTML = System.IO.File.ReadAllText(htmlFilePath.ToString());

            //Loop and replace the Image Path.
            foreach (Match match in Regex.Matches(wordHTML, "<v:imagedata.+?src=[\"'](.+?)[\"'].*?>", RegexOptions.IgnoreCase))
            {
                wordHTML = Regex.Replace(wordHTML, match.Groups[1].Value, "Temp/" + match.Groups[1].Value);
            }

            //Delete the Uploaded Word File.
            System.IO.File.Delete(fileSavePath.ToString());

            ViewBag.WordHtml = wordHTML;

            return View();
        }



        //Code to read word table data
        public ActionResult ReadwordTable()
        {
            return View();
        }
        public ActionResult Readtabledata(HttpPostedFileBase uploadfile)
        {
            if (uploadfile != null)
            {
                ReadDataTableModel rdtm = new ReadDataTableModel();
                //try
                //{
                Process[] _proceses = null;
                _proceses = Process.GetProcessesByName("WINWORD");
                foreach (Process proces in _proceses)
                {
                    proces.Kill();
                }
                    Application application = new Application();
                    
                    
                    //Microsoft.Office.Interop.Word.Document doc = application.Documents.Open(Environment.CurrentDirectory + "\\Functions.docx", ReadOnly: false, Visible: false);
                    uploadfile.SaveAs(Server.MapPath("\\Resume\\") + uploadfile.FileName);
                    string filepath = Server.MapPath("\\Resume\\" + uploadfile.FileName.ToString());

                    Microsoft.Office.Interop.Word.Document doc = application.Documents.Open(filepath, ReadOnly: false, Visible: true);
                    Microsoft.Office.Interop.Word.Table table = doc.Tables[1];





                    rdtm.name = table.Cell(2, 2).Range.Text.ToString();
                    rdtm.cnt = table.Cell(2, 4).Range.Text.ToString();
                    rdtm.qualification = table.Cell(3, 2).Range.Text.ToString();
                    rdtm.email = table.Cell(3, 4).Range.Text.ToString();

                    //table.Cell(4, 4).Range.Text = "someString";
                    
                    doc.Save();
                    doc = null;
                    
                    
                //}
                //catch (IOException ioe)
                //{
                //    var msg = ioe.Message.ToString();
                    
                //}
                return View(rdtm);
            }
            else
            {
                return View();
            }
            
        }
    }
}