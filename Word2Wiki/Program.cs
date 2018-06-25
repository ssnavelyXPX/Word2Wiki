using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Word2Wiki
{
    class Program
    {
        static void Main(string[] args)
        {
            StringBuilder sb = new StringBuilder();
            DocumentObject.LoadDocument(@"C:\Users\ssnavely\Downloads\testing.docx");
            DocumentObject.Doc.Paragraphs.ToList().ForEach(p => sb.Append(new WikiParagraph(p)).Append(Environment.NewLine));

            File.WriteAllText(@"C:\Users\ssnavely\Documents\wikiTest.txt", sb.ToString());

            PrintXmlToConsole(); 
        }

        static void PrintXmlToConsole()
        {
            DocumentObject.Doc.Paragraphs.ToList().ForEach(p => Console.WriteLine(p.Xml));
            Console.ReadLine();
        }
    }
}
