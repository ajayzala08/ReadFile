using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Code7248.word_reader;

namespace ReadFile
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
            string path = @"F:\Download backup 7th august 2018\Ajay Resume.docx";
            cs.readFileContent(path);
            Console.ReadLine();
            //code for github desktop application
        }
    }
}
