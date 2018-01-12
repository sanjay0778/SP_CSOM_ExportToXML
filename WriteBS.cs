using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExportToXML
{
    public class WriteSB
    {
        public static void CreateMfmatrNode(XmlDocument xmlDoc, XmlNode nodeMFMATR)
        {
            //TestIncSheet
            XmlNode nodeIncSheet = xmlDoc.CreateElement("MFITEM");
            XmlAttribute attIncSheetType = xmlDoc.CreateAttribute("TYPE");
            attIncSheetType.InnerText = "FILE";
            nodeIncSheet.Attributes.Append(attIncSheetType);
            XmlAttribute attIncSheetKey = xmlDoc.CreateAttribute("KEY");
            attIncSheetKey.InnerText = "TestIncSheet.pdf";
            nodeIncSheet.Attributes.Append(attIncSheetKey);
            XmlNode nodeIncSheetTitle = xmlDoc.CreateElement("TITLE");
            nodeIncSheetTitle.InnerText = "<![CDATA[TestInc Sheet]]>";
            nodeIncSheet.AppendChild(nodeIncSheetTitle);
            nodeMFMATR.AppendChild(nodeIncSheet);
        }

        public static XmlNode DocumentSetNodeSBLevel6(ListItem item, XmlDocument xmlDoc, string name, string strDescription)
        {
            //Level 6 for document sets
            XmlNode nodeLevel6 = xmlDoc.CreateElement("LEVEL6");
            XmlAttribute attLevel6Key = xmlDoc.CreateAttribute("KEY");
            XmlAttribute attLevel6Type = xmlDoc.CreateAttribute("TYPE");
            XmlNode nodeLevel6Title = xmlDoc.CreateElement("TITLE");

            // Level 6 element
            attLevel6Type.InnerText = "TAXONOMY";
            nodeLevel6Title.InnerText = "<![CDATA[" + name + " :: " + strDescription + "]]>";
            attLevel6Key.InnerText = name;
            nodeLevel6.Attributes.Append(attLevel6Key);
            nodeLevel6.Attributes.Append(attLevel6Type);
            nodeLevel6.AppendChild(nodeLevel6Title);

            return nodeLevel6;
        }

        public static XmlNode DocumentNodeBS(ListItem item, XmlDocument xmlDoc, DateTime inputDate)
        {
            string name = string.Empty;
            string displayName = string.Empty;
            //string description = string.Empty;
            DateTime modifiedDate = DateTime.MinValue;
            DateTime createdDate = DateTime.MinValue;
            //Level 6 for documents
            XmlNode nodeLevel7 = xmlDoc.CreateElement("LEVEL7");
            XmlAttribute attLevel7Key = xmlDoc.CreateAttribute("KEY");
            XmlAttribute attLevel7Type = xmlDoc.CreateAttribute("TYPE");
            XmlAttribute attLevel7CHG = xmlDoc.CreateAttribute("CHG");
            XmlNode nodeLevel7Title = xmlDoc.CreateElement("TITLE");

            if (item != null)
            {
                name = Convert.ToString(item["FileLeafRef"]);
                displayName = item.DisplayName;
                //description = Convert.ToString(item["ProductDescription"]);
                modifiedDate = Convert.ToDateTime(item["Modified"]);
                createdDate = Convert.ToDateTime(item["Created"]);
            }

            //level 6 element
            attLevel7Type.InnerText = "FILE";
            nodeLevel7Title.InnerText = "<![CDATA[" + displayName + "]]>";
            attLevel7Key.InnerText = name;
            // Logic for xxx tag
            TimeSpan span = modifiedDate.Subtract(createdDate);
            double diff = span.TotalSeconds;

            if (diff < 10) // assuming document is not modified within 10 sec.               
            {
                if (createdDate >= inputDate)
                {
                    attLevel7CHG.InnerText = "N";//created date is the same or later than the user specified date 
                }
                else
                {
                    attLevel7CHG.InnerText = "U"; //Confirmed with Felice
                }
            }
            else
            {
                if (modifiedDate >= inputDate)
                {
                    attLevel7CHG.InnerText = "R";//modified date is the same or later than the user specified date
                }
                else
                {
                    attLevel7CHG.InnerText = "U";
                }
            }
            nodeLevel7.Attributes.Append(attLevel7Key);
            nodeLevel7.Attributes.Append(attLevel7Type);
            nodeLevel7.Attributes.Append(attLevel7CHG);
            nodeLevel7.AppendChild(nodeLevel7Title);
            return nodeLevel7;
        }
    }
}
