using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExportToXML
{
    class RepositoryData
    {
        Logger logger = new Logger();
        public ListItemCollection GetDocMetaData(ClientContext clientContext, string libraryName, string contentTypeName, string xxxNumber)
        {
            var ids = GetxxxNumberList(xxxNumber);

            ListItemCollection items = null;
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;

                List library = clientContext.Web.Lists.GetByTitle(libraryName);

                clientContext.Load(library);
                clientContext.Load(library.RootFolder);
                clientContext.Load(library.RootFolder.Folders);
                clientContext.Load(library.RootFolder.Files);
                clientContext.ExecuteQuery();

                if (library != null && library.ItemCount > 0)
                {
                    List<string> conditions = new List<string>();
                    foreach (string id in ids)
                    {
                        conditions.Add("<Eq><FieldRef Name='xxx' /><Value Type='Text'>" + id + "</Value></Eq>");
                    }
                    string merged = MergeCAMLConditions(conditions, MergeType.Or);                    
                    CamlQuery camlQuery = new CamlQuery();
                    string serverRelativeUrl = library.RootFolder.ServerRelativeUrl + "/SYZ Document";
                    camlQuery.FolderServerRelativeUrl = serverRelativeUrl;
                    string viewXml = @"<View Scope='Recursive'>
                                            <Query><Where>" + merged + 
                                       @"</Where></Query> 
                                                                                       
                                        </View>";
                    camlQuery.ViewXml = viewXml;

                    items = library.GetItems(camlQuery);

                    clientContext.Load(items, listItems => listItems.Include
                                            (item => item.FileSystemObjectType,
                                             item => item.ContentType,
                                             item => item.DisplayName,                                             
                                             item => item["FileLeafRef"],
                                             item => item["Modified"],
                                             item => item["Created"],
                                             item => item["Title"],
                                             item => item["FileRef"],
                                             item => item["ProductDescription"],
                                             item => item["Confidentiality"],
                                             item => item["Created"]));
                    clientContext.ExecuteQuery();

                }

            }
            catch (Exception exec)
            {
                logger.WriteEventLog("Please contact admin : " + exec.Message);
                throw exec;
            }
            int count = items.Count;
            return items;
        }

        /// <summary>
        /// This method loads xxxs in list box from xxx list in SharePoint
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        public ListItemCollection LoadxxxList(ClientContext clientContext)
        {
            Web site = clientContext.Web;
            List oList = clientContext.Web.Lists.GetByTitle("xxx");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = @"<View Scope='Recursive'>
                                            <Query></Query>
                                        </View>";
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem, listItems => listItems.Include
                                           (item => item["Col1"],
                                            item => item["Col2"]));

            clientContext.ExecuteQuery();
            return collListItem;
        }

        enum MergeType { Or, And };
        /// <summary>
        /// For Dynamic CAML Query
        /// </summary>
        /// <param name="conditions"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private string MergeCAMLConditions(List<string> conditions, MergeType type)
        {
            try
            {
                if (conditions.Count == 0) return "";
                string typeStart = (type == MergeType.And ? "<And>" : "<Or>"); string typeEnd = (type == MergeType.And ? "</And>" : "</Or>");
                // Build hierarchical structure   
                while (conditions.Count >= 2)
                {
                    List<string> complexConditions = new List<string>();
                    for (int i = 0; i < conditions.Count; i += 2)
                    {
                        if (conditions.Count == i + 1)
                            // Only one condition left               
                            complexConditions.Add(conditions[i]);
                        else
                            // Two condotions - merge               
                            complexConditions.Add(typeStart + conditions[i] + conditions[i + 1] + typeEnd);
                    }
                    conditions = complexConditions;
                }
            }
            catch (Exception exec)
            {
                logger.WriteEventLog("Please contact admin : " + exec.Message);
                throw exec;
            }
            return conditions[0];
        }

        /// <summary>
        /// Returns list of xxx numbers
        /// </summary>
        /// <param name="strxxxNumber"></param>
        /// <returns></returns>
        public List<string> GetxxxNumberList(string strxxxNumber)
        {
            strxxxNumber = strxxxNumber.Replace('\n', ',');
            if (strxxxNumber.EndsWith((",")))
            {
                strxxxNumber = strxxxNumber.Substring(0, strxxxNumber.Length - 1);
            }
            string[] result = strxxxNumber.Split(',');
            List<string> strResult = new List<string>();
            for (int i = 0; i < result.Length; i++)
            {

                strResult.Add(result[i]);
            }
            return strResult;
        }

    }
}
