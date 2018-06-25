using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Word2Wiki
{
    public class Run : DocumentObject
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string InnerText { get; set; }
          

        public Run(XmlNode node)
        {
            XmlNode runProp = node.SelectSingleNode("./w:rPr", manager);

            if(runProp != null)
            {
                Bold = runProp.SelectSingleNode("./w:b", manager) != null;
                Italic = runProp.SelectSingleNode("./w:i", manager) != null;
                
            } else
            {
                Bold = false;
                Italic = false;
            }
            InnerText = node.InnerText;
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            if (Bold && Italic)
            {
                sb.Append("<b><i>");
                sb.Append(InnerText);
                sb.Append("</b></i>");
            } else if (Bold)
            {
                sb.Append("<b>");
                sb.Append(InnerText);
                sb.Append("</b>");
            }
            else if (Italic)
            {
                sb.Append("<i>");
                sb.Append(InnerText);
                sb.Append("</i>");
            }
            else
            {
                sb.Append(InnerText);
            }

            return sb.ToString();
        }
    }
}