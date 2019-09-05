using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WebApplication4.Models;
using System.IO; 

namespace WebApplication4.Controllers
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
        public ActionResult ReadFile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Readwordfile(HttpPostedFileBase uploadfile)
        {
            if (uploadfile != null)
            {
                FileInfo fi = new FileInfo(uploadfile.FileName.ToString());
                if (fi.Extension.ToLower() == ".docx")
                {

                    tble_model tmodel = new tble_model();
                    uploadfile.SaveAs(Server.MapPath("\\Resume\\") + uploadfile.FileName);
                    string fileName = Server.MapPath("\\Resume\\" + uploadfile.FileName.ToString());

                    // string fileName = @"F:\Users\Arche\Desktop\Binesh KK 1978 BTech(Electro)+MBA(Operation) Asst Mgr(Strategic Sourcing) Bangalore - good.docx";
                    // The text to be added to the cell (3,2).  
                    string addedText = "This is the text added by the API example";
                    // Open the file for editing.  
                    using (WordprocessingDocument doc =
                        WordprocessingDocument.Open(fileName, true))
                    {
                        // Find the first table in the document.  
                        Table table =
                            doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                        // Find the second row in the table.  
                        TableRow row = table.Elements<TableRow>().ElementAt(1);

                        // Find the second cell in the row.  
                        //TableCell cell = row.Elements<TableCell>().ElementAt(1);
                        tmodel.name = row.Elements<TableCell>().ElementAt(1).InnerText.ToString();
                        tmodel.cnt = row.Elements<TableCell>().ElementAt(3).InnerText.ToString();


                        // Find the third row in the table
                        TableRow row1 = table.Elements<TableRow>().ElementAt(2);
                        tmodel.qualification = row1.Elements<TableCell>().ElementAt(1).InnerText.ToString();
                        tmodel.email = row1.Elements<TableCell>().ElementAt(3).InnerText.ToString();

                        // Find the forth row in the table
                        TableRow row2 = table.Elements<TableRow>().ElementAt(3);
                        tmodel.company = row2.Elements<TableCell>().ElementAt(1).InnerText.ToString();



                        // Find the fifth row in the table
                        TableRow row3 = table.Elements<TableRow>().ElementAt(4);
                        tmodel.designation = row3.Elements<TableCell>().ElementAt(1).InnerText.ToString();


                        // Find the sixth row in the table
                        TableRow row4 = table.Elements<TableRow>().ElementAt(5);
                        tmodel.location = row4.Elements<TableCell>().ElementAt(1).InnerText.ToString();





                        // Find the first paragraph in the table cell.  
                        //Paragraph parag = cell.Elements<Paragraph>().First();

                        // Find the first run in the paragraph.  
                        //Run run = parag.Elements<Run>().First();

                        // Set the text for the run.  
                        //Text text = run.Elements<Text>().First();
                        //text.Text = addedText;
                    }

                    // Console.WriteLine("All done. Press any key.");
                    // Console.ReadKey(); 
                    return View(tmodel);
                }
                else
                {
                    return RedirectToAction("ReadFile");
                }
                
            }
            else
            {
                return RedirectToAction("ReadFile");
            }
            
        }
    }
}