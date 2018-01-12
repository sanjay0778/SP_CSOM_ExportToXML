using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ExportToXML
{
    class WriteXML
    {
        Logger logger = new Logger();
        FormatXML formatXml = new FormatXML();
        DownloadFile download = new DownloadFile();
        public void CreateXml(ListItemCollection InputList, DateTime date, string docType, List<ModXXFamXX> ModXXs, ClientContext clientContext)
        {
            try
            {
                DateTime inputDate = date;
                XmlDocument xmlDoc = new XmlDocument();
                string itemType = string.Empty; // to check item is document set(folder) or document(file)
                string contentType = string.Empty; //to get content type
                string confidentiality = string.Empty;

                XmlNode nodeLevel4 = StaticNodes(xmlDoc, docType);
                XmlNode nodeLevel5Dynamic = null;
                XmlNode nodeLevel6Dynamic = null;

                //Families for selected ModXXs
                var selectedFamXX = (from A in ModXXs
                                      select A.FamilAlias).ToList().Distinct().ToList();
                //Selected ModXXs
                var selectedModXX = (from A in ModXXs
                                     select A.ModXXName).ToList();

                var listItems = InputList.ToList();
                if (docType == "WAR")
                {
                    listItems = InputList.OrderBy(c => c["TATNS"]).ThenByDescending(n => n.FileSystemObjectType).ThenBy(n => n["FileLeafRef"]).ToList();
                }
                else
                {
                    listItems = InputList.OrderBy(c => c["TATNS"]).ThenByDescending(n => n.FileSystemObjectType).ThenBy(n => n["FileLeafRef"]).ToList();
                }

                for (int i = 0; i < listItems.Count; i++)
                {
                    itemType = listItems[i].FileSystemObjectType.ToString();
                    contentType = listItems[i].ContentType.Name;                   
                    confidentiality = Convert.ToString(listItems[i]["Confidentiality"]);
                    if (confidentiality == "Public" || String.IsNullOrEmpty(confidentiality))
                    {

                        if (docType == "WAR")
                        {
                            if (contentType.Contains("My ContentType"))
                            {
                                if (itemType == "Folder")
                                {
                                    string strTATNS = string.Empty;
                                    strTATNS = Convert.ToString(listItems[i]["TATNS"]);
                                    string tatNumber = strTATNS.Split(new char[0])[0];
                                    // Get list of existing LEVEL5 nodes.
                                    string level5 = String.Format("/LEVEL3/LEVEL4/LEVEL5[@KEY='{0}']", tatNumber);
                                    XmlNodeList node5List = xmlDoc.SelectNodes(level5);
                                    if (node5List.Count > 0)
                                    {
                                        nodeLevel5Dynamic = node5List[0];
                                    }
                                    else
                                    {
                                        nodeLevel5Dynamic = DocumentSetNode(listItems[i], xmlDoc, tatNumber, strTATNS);
                                    }
                                    nodeLevel4.AppendChild(nodeLevel5Dynamic);
                                }
                                else
                                {
                                    // check needs to be added for files with name starting with WAR
                                    if (listItems[i]["FileLeafRef"].ToString().StartsWith("WAR"))
                                    {
                                        XmlNode nodeLevel6DynamicAw = DocumentNode(listItems[i], xmlDoc, inputDate);
                                        nodeLevel5Dynamic.AppendChild(nodeLevel6DynamicAw);

                                    }
                                }
                            }
                        }
                        else
                        {
                            if (contentType.Contains("MyBulletin"))
                            {
                                string strTATNS = string.Empty;
                                string strDescription = string.Empty;
                                string name = string.Empty;
                                strTATNS = Convert.ToString(listItems[i]["TATNS"]);
                                string tatNumber = strTATNS.Split(new char[0])[0];
                                name = Convert.ToString(listItems[i]["FileLeafRef"]);
                                strDescription = Convert.ToString(listItems[i]["ProductDescription"]);

                                if (itemType == "Folder")
                                {
                                    // Get list of existing LEVEL5 nodes.
                                    string level5 = String.Format("/LEVEL3/LEVEL4/LEVEL5[@KEY='{0}']", tatNumber);
                                    XmlNodeList node5List = xmlDoc.SelectNodes(level5);

                                    if (node5List.Count > 0)
                                    {
                                        nodeLevel5Dynamic = node5List[0];
                                    }
                                    else
                                    {
                                        nodeLevel5Dynamic = DocumentSetNode(listItems[i], xmlDoc, tatNumber, strTATNS);
                                    }
                                    nodeLevel4.AppendChild(nodeLevel5Dynamic);
                                    // Get list of existing LEVEL6 nodes.
                                    string level6 = String.Format("/LEVEL3/LEVEL4/LEVEL5/LEVEL6[@KEY='{0}']", name);
                                    XmlNodeList node6List = xmlDoc.SelectNodes(level6);
                                    if (node6List.Count > 0)
                                    {
                                        nodeLevel6Dynamic = node6List[0];
                                    }
                                    else
                                    {
                                        nodeLevel6Dynamic = WriteSB.DocumentSetNodeSBLevel6(listItems[i], xmlDoc, name, strDescription);
                                    }
                                    nodeLevel5Dynamic.AppendChild(nodeLevel6Dynamic);
                                }
                                else
                                {
                                    // Find level6 node with same title as level 7 then add. Find level6 in all items with same title as listItems[i]
                                    string title = Convert.ToString(listItems[i]["Title"]);
                                    var docSetWithSameTitle = listItems.Where(s => Convert.ToString(s["Title"]) == title).Where(s => s.FileSystemObjectType == FileSystemObjectType.Folder).ToList();
                                    //Check if node 6 has same name as docset above.
                                    if (docSetWithSameTitle.Count > 0)
                                    {
                                        string nameofDocSet = Convert.ToString(docSetWithSameTitle[0]["FileLeafRef"]);
                                        string level6 = String.Format("/LEVEL3/LEVEL4/LEVEL5/LEVEL6[@KEY='{0}']", nameofDocSet);
                                        XmlNodeList node6List = xmlDoc.SelectNodes(level6);
                                        if (node6List.Count > 0)
                                        {
                                            nodeLevel6Dynamic = node6List[0];
                                        }

                                        XmlNode nodeLevel7Dynamic = WriteSB.DocumentNodeSB(listItems[i], xmlDoc, inputDate);
                                        nodeLevel6Dynamic.AppendChild(nodeLevel7Dynamic);
                                    }
                                    else
                                    {
                                        logger.LogMissingTitles("Following file is missing Title. Downloaded but not included in XML: " + listItems[i].DisplayName, selectedFamXX, selectedModXX);
                                        //logger.WriteEventLog("Following file is missing Title. Downloaded but not included in XML: " + listItems[i].DisplayName);
                                    }
                                    // Download file
                                    download.FileRef(listItems[i], clientContext, selectedFamXX, selectedModXX);
                                }
                            }
                        }
                    }
                }

                string fileDirectory = ConfigurationManager.AppSettings["XmlFileLocation"];
                if (!Directory.Exists(fileDirectory))
                {
                    Directory.CreateDirectory(fileDirectory);
                }

                if (docType == "WAR")
                {
                    string strAbbreviation = string.Join("", selectedFamXX.ToArray());
                    string strModXX = string.Join("", selectedModXX.ToArray()) + "_AW_INDEX";
                    string xmlFileLocation = fileDirectory + strAbbreviation + strModXX + ".xml";
                    xmlDoc.Save(xmlFileLocation);
                    string formattedString = formatXml.ReplaceSpecialChar(xmlFileLocation);
                    XmlDocument formattedXml = new XmlDocument();
                    formattedXml.LoadXml(formattedString);
                    formattedXml.Save(xmlFileLocation);
                }
                else
                {
                    string strSBAbbreviation = string.Join("", selectedFamXX.ToArray());
                    string strModXX = string.Join("", selectedModXX.ToArray()) + "_SB_INDEX";
                    string xmlFileLocation = fileDirectory + strSBAbbreviation + strModXX + ".xml";
                    xmlDoc.Save(xmlFileLocation);
                    string formattedString = formatXml.ReplaceSpecialChar(xmlFileLocation);
                    XmlDocument formattedXml = new XmlDocument();
                    formattedXml.LoadXml(formattedString);
                    formattedXml.Save(xmlFileLocation);
                }
            }
            catch (Exception exec)
            {
                logger.WriteEventLog("Please contact admin : " + exec.Message);
                throw exec;
            }
        }

        private static XmlNode StaticNodes(XmlDocument xmlDoc, string docType)
        {
            #region "Static Values"
            XmlNode nodeDec = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            xmlDoc.AppendChild(nodeDec);

            //Level3 - RootNode creation constant
            XmlNode nodeLevel3 = xmlDoc.CreateElement("LEVEL3");
            XmlAttribute attType = xmlDoc.CreateAttribute("TYPE");
            attType.InnerText = "ABC";
            nodeLevel3.Attributes.Append(attType);
            XmlAttribute attKey = xmlDoc.CreateAttribute("KEY");
            if (docType == "WAR")
            {
                attKey.InnerText = "WAR";
            }
            else
            {
                attKey.InnerText = "SA";
            }

            nodeLevel3.Attributes.Append(attKey);
            xmlDoc.AppendChild(nodeLevel3);

            //TSN creation constant
            XmlNode nodeTSN = xmlDoc.CreateElement("TSN");
            nodeTSN.InnerText = "-1";
            nodeLevel3.AppendChild(nodeTSN);

            //REVDATE creation constant
            XmlNode nodeREVDATE = xmlDoc.CreateElement("REVDATE");
            nodeREVDATE.InnerText = DateTime.Now.ToString("yyyyMMdd");
            nodeLevel3.AppendChild(nodeREVDATE);

            //MFMATR creation constant -for WAR and SA both
            XmlNode nodeMFMATR = xmlDoc.CreateElement("MFMATR");
            XmlAttribute attMFMType = xmlDoc.CreateAttribute("TYPE");
            attMFMType.InnerText = "FOLDERCPT";
            nodeMFMATR.Attributes.Append(attMFMType);
            XmlAttribute attMFMKey = xmlDoc.CreateAttribute("KEY");
            attMFMKey.InnerText = "MFMATR";
            nodeMFMATR.Attributes.Append(attMFMKey);
            nodeLevel3.AppendChild(nodeMFMATR);

            //MFITEM creation constant - for WAR
            if (docType == "WAR")
            {
                XmlNode nodeMFITEM = xmlDoc.CreateElement("MFITEM");
                XmlAttribute attMFTYPE = xmlDoc.CreateAttribute("TYPE");
                attMFTYPE.InnerText = "FILE";
                nodeMFITEM.Attributes.Append(attMFTYPE);
                XmlAttribute attMFKEY = xmlDoc.CreateAttribute("KEY");
                attMFKEY.InnerText = "USERCOMM.full.pdf";
                nodeMFITEM.Attributes.Append(attMFKEY);
                XmlNode nodeMFTitle = xmlDoc.CreateElement("TITLE");
                nodeMFTitle.InnerText = "<![CDATA[User Comments]]>";
                nodeMFITEM.AppendChild(nodeMFTitle);
                nodeMFMATR.AppendChild(nodeMFITEM);
            }
            else
            {
                WriteSB.CreateMfmatrNode(xmlDoc, nodeMFMATR);
            }

            //Level4 creation constant
            XmlNode nodeLevel4 = xmlDoc.CreateElement("LEVEL4");
            XmlAttribute attLevel4Type = xmlDoc.CreateAttribute("TYPE");
            attLevel4Type.InnerText = "FOLDERCPT";
            nodeLevel4.Attributes.Append(attLevel4Type);
            XmlAttribute attLevel4Key = xmlDoc.CreateAttribute("KEY");
            attLevel4Key.InnerText = "99";
            nodeLevel4.Attributes.Append(attLevel4Key);
            XmlNode nodeLevel4Title = xmlDoc.CreateElement("TITLE");
            if (docType == "WAR")
            {
                nodeLevel4Title.InnerText = "<![CDATA[My Wires (Updated on " + DateTime.Now.ToString("MMM dd/yyyy") + ")]]>";
            }
            else
            {
                nodeLevel4Title.InnerText = "<![CDATA[My Bulletins (Updated on " + DateTime.Now.ToString("MMM dd/yyyy") + ")]]>";
            }
            nodeLevel4.AppendChild(nodeLevel4Title);
            nodeLevel3.AppendChild(nodeLevel4);

            //Level5 creation constant - only for WAR
            if (docType == "WAR")
            {
                XmlNode nodeLevel5 = xmlDoc.CreateElement("LEVEL5");
                XmlAttribute attLevel5Key = xmlDoc.CreateAttribute("KEY");
                attLevel5Key.InnerText = "Index";
                XmlAttribute attLevel5Type = xmlDoc.CreateAttribute("TYPE");
                attLevel5Type.InnerText = "TAXONOMY";
                XmlNode nodeLevel5Title = xmlDoc.CreateElement("TITLE");
                nodeLevel5Title.InnerText = "<![CDATA[Index]]>";
                nodeLevel5.AppendChild(nodeLevel5Title);

                nodeLevel5.Attributes.Append(attLevel5Key);
                nodeLevel5.Attributes.Append(attLevel5Type);

                //Level6 creation constant  
                XmlNode nodeLevel6 = xmlDoc.CreateElement("LEVEL6");
                XmlAttribute attLevel6Key = xmlDoc.CreateAttribute("KEY");
                attLevel6Key.InnerText = "WAR_INDEX.pdf";
                XmlAttribute attLevel6Type = xmlDoc.CreateAttribute("TYPE");
                attLevel6Type.InnerText = "FILE";
                XmlNode nodeLevel6Title = xmlDoc.CreateElement("TITLE");
                nodeLevel6Title.InnerText = "<![CDATA[by WAR#, ATA, Date, Subject]]>";
                XmlAttribute attLevel6CHG = xmlDoc.CreateAttribute("CHG");
                attLevel6CHG.InnerText = "R";
                nodeLevel6.Attributes.Append(attLevel6Type);
                nodeLevel6.Attributes.Append(attLevel6Key);
                nodeLevel6.Attributes.Append(attLevel6CHG);
                nodeLevel6.AppendChild(nodeLevel6Title);
                nodeLevel5.AppendChild(nodeLevel6);

                nodeLevel4.AppendChild(nodeLevel5);
            }

            #endregion
            return nodeLevel4;
        }

        private XmlNode DocumentSetNode(ListItem item, XmlDocument xmlDoc, string tatNumber, string TATNSNumber)
        {
            //Level 5 for document sets
            XmlNode nodeLevel5 = xmlDoc.CreateElement("LEVEL5");
            XmlAttribute attLevel5Key = xmlDoc.CreateAttribute("KEY");
            XmlAttribute attLevel5Type = xmlDoc.CreateAttribute("TYPE");
            XmlNode nodeLevel5Title = xmlDoc.CreateElement("TITLE");

            // Level 5 element
            attLevel5Type.InnerText = "TAXONOMY";
            nodeLevel5Title.InnerText = String.Format("<![CDATA[{0}]]>", TATNSNumber);
            attLevel5Key.InnerText = tatNumber;
            nodeLevel5.Attributes.Append(attLevel5Key);
            nodeLevel5.Attributes.Append(attLevel5Type);
            nodeLevel5.AppendChild(nodeLevel5Title);
            return nodeLevel5;
        }

        private XmlNode DocumentNode(ListItem item, XmlDocument xmlDoc, DateTime inputDate)
        {
            string strTitle = string.Empty;
            string name = string.Empty;
            string description = string.Empty;
            DateTime modifiedDate = DateTime.MinValue;
            DateTime createdDate = DateTime.MinValue;
            //Level 6 for documents
            XmlNode nodeLevel6 = xmlDoc.CreateElement("LEVEL6");
            XmlAttribute attLevel6Key = xmlDoc.CreateAttribute("KEY");
            XmlAttribute attLevel6Type = xmlDoc.CreateAttribute("TYPE");
            XmlAttribute attLevel6CHG = xmlDoc.CreateAttribute("CHG");
            XmlNode nodeLevel6Title = xmlDoc.CreateElement("TITLE");

            if (item != null)
            {
                strTitle = Convert.ToString(item["Title"]);
                name = Convert.ToString(item["FileLeafRef"]);
                description = Convert.ToString(item["ProductDescription"]);
                modifiedDate = Convert.ToDateTime(item["Modified"]);
                createdDate = Convert.ToDateTime(item["Created"]);
            }

            //level 6 element
            attLevel6Type.InnerText = "FILE";
            nodeLevel6Title.InnerText = "<![CDATA[" + strTitle + " :: " + description + "]]>";
            attLevel6Key.InnerText = "_Attachments/" + name;
            // Logic for CHG tag
            TimeSpan span = modifiedDate.Subtract(createdDate);
            double diff = span.TotalSeconds;

            if (diff < 10) // assuming document is not modified within 10 sec.               
            {
                if (createdDate >= inputDate)
                {
                    attLevel6CHG.InnerText = "N";//created date is the same or later than the user specified date 
                }
                else
                {
                    attLevel6CHG.InnerText = "U"; //Confirmed with Felice
                }
            }
            else
            {
                if (modifiedDate >= inputDate)
                {
                    attLevel6CHG.InnerText = "R";//modified date is the same or later than the user specified date
                }
                else
                {
                    attLevel6CHG.InnerText = "U";
                }
            }
            nodeLevel6.Attributes.Append(attLevel6Key);
            nodeLevel6.Attributes.Append(attLevel6Type);
            nodeLevel6.Attributes.Append(attLevel6CHG);
            nodeLevel6.AppendChild(nodeLevel6Title);
            return nodeLevel6;
        }

    }
}
