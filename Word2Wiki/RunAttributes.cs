using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Word2Wiki
{
    class RunAttributes : DocumentObject
    {
        public RunAttributes(XmlNode runProp)
        {
            if (runProp != null)
            {
                if (runProp.SelectSingleNode("./w:rStyle[@w:val = 'Hyperlink']", manager) != null)
                {
                    string rId = runProp.SelectSingleNode("../../@r:id", manager).Value;
                    PackageRelationship rel = Doc.PackagePart.GetRelationships().Single(
                        r => r.Id.Equals(rId));
                    Console.WriteLine(rel.TargetUri);

                }
                Bold = runProp.SelectSingleNode("./w:b", manager) != null;
                Italic = runProp.SelectSingleNode("./w:i", manager) != null;

            }
            else
            {
                Bold = false;
                Italic = false;
                //Hyperlink = null;
            }
        }

        public bool Bold { get; set; }
        public bool Italic { get; set; }
        //public Hyperlink Hyperlink { get; set; }
    }
}
