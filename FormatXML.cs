using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;


namespace ExportToXML
{
    public class FormatXML
    {
        /// <summary>
        /// Read the xml file to string
        /// </summary>
        /// <param name="fileName">full file path</param>
        /// <returns></returns>
        public string ReadXmlFile(string fileName)
        {
            string xmlString = System.IO.File.ReadAllText(fileName);
            return xmlString;
        }

        public string ReplaceSpecialChar(string fileName)
        {
            string xmlString = ReadXmlFile(fileName);
            xmlString = xmlString.Replace("&lt;", "<").Replace("&gt;", ">");
            return xmlString;
        }
    }
}
