using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace DocXFunc.ex
{
    public static class XceedExtensions
    {
        public static XElement AddElement(this XElement parent, string name)
        {
            var element = new XElement(XName.Get(name, "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
            parent.Add(element);
            return element;
        }
    }
}
