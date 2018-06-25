using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Xceed.Words.NET;

namespace Word2Wiki
{
    
    public class DocumentObject
    {
        public static DocX Doc;
        public static XmlNamespaceManager manager;
        public static XmlDocument Xml;

        public static void LoadDocument(string path)
        {
            Doc = DocX.Load(path);
            Xml = new XmlDocument();
            Xml.LoadXml(Doc.Xml.ToString());

            manager = new XmlNamespaceManager(Xml.NameTable);
            manager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            manager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        }

    }
}
