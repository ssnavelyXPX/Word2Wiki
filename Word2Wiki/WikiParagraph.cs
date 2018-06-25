using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Xceed.Words.NET;

namespace Word2Wiki
{
    public class WikiParagraph : DocumentObject
    {
        private readonly Paragraph paragraph;
        private readonly StringBuilder sb = new StringBuilder();

        public WikiParagraph(Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }

        public override string ToString()
        {
            switch (paragraph.StyleName)
            {
                case "Heading1":
                    sb.Append("==");
                    GetRuns().ForEach(run => sb.Append(run));
                    sb.Append("==");
                    break;
                case "Heading2":
                    sb.Append("===");
                    GetRuns().ForEach(run => sb.Append(run));
                    sb.Append("===");
                    break;
                case "Heading3":
                    sb.Append("====");
                    GetRuns().ForEach(run => sb.Append(run));
                    sb.Append("====");
                    break;
                case "ListParagraph":
                    sb.Append(GetListPrefixAsString());
                    GetRuns().ForEach(run => sb.Append(run));
                    break;
                case "Normal":
                    GetRuns().ForEach(run => sb.Append(run));
                    break;
            }

            return sb.ToString();
        }

        public List<Run> GetRuns()
        {
            List<Run> runs = new List<Run>();

            XmlDocument paragraphXml = new XmlDocument();
            paragraphXml.LoadXml(paragraph.Xml.ToString());

            foreach (XmlNode node in paragraphXml.SelectNodes("//w:r", manager))
            {
                runs.Add(new Run(node));
            }

            return runs;
        }

        private string GetListPrefixAsString()
        {
            string bulletType = "";
            int indent = (paragraph.IndentLevel ?? 0) + 1;

            switch (paragraph.ListItemType)
            {
                case ListItemType.Bulleted:
                    bulletType = "*";
                    break;
                case ListItemType.Numbered:
                    bulletType = "#";
                    break;
            }

            return String.Concat(Enumerable.Repeat(bulletType, indent));
        }
    }
}