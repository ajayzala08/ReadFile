using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office;
using Microsoft.Office.Interop.Word;



namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            //WordFileToRead.SaveAs(Server.MapPath(WordFileToRead.FileName));
            object filename = @"F:\Users\Arche\Desktop\Project.docx";//Server.MapPath(WordFileToRead.FileName);
            ApplicationClass AC = new ApplicationClass();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;
            try
            {
                doc = AC.Documents.Open(ref filename, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref isVisible, ref missing, ref missing, ref missing);
                WordFileText.Text = doc.Content.Text;
            }
            catch (Exception ex)
            {

            }  
        }
    }
}
