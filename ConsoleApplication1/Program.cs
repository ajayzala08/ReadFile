using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Code7248.word_reader;

namespace ConsoleApplication1
{
  
        class Program
        {
            private void readFileContent(string path)
            {
                TextExtractor extractor = new TextExtractor(path);
                string text = extractor.ExtractText();
                Console.WriteLine(text);
            }
            static void Main(string[] args)
            {
                Program cs = new Program();
                string path = @"F:\Users\Arche\Desktop\Project.docx";
                cs.readFileContent(path);
                Console.ReadLine();
            }
        }
    
}
