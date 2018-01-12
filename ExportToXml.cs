using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentSubclass;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExportToXML
{
    public partial class ExportToXml : System.Windows.Forms.Form
    {
        WriteXML writeXml = null;
        RepositoryData repData = null;
        Logger logger = null;
        ListItemCollection collListItem = null;
        List<SubclassClass> lstMF = new List<SubclassClass>();
        public ExportToXml()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Load event of windows form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportToXml_Load(object sender, EventArgs e)
        {

            try
            {
                lblMessage.Visible = false;
                cbDocType.SelectedIndex = 0;
                repData = new RepositoryData();

                string siteUrl = ConfigurationManager.AppSettings["Site"];
                ClientContext clientContext = new ClientContext(siteUrl);
                collListItem = repData.LoadSubclassList(clientContext);
                lstMF = new List<SubclassClass>();

                foreach (ListItem oListItem in collListItem)
                {
                    if (oListItem["Subclass"].ToString() != "All")
                    {
                        FieldLookupValue lookup = oListItem["Class"] as FieldLookupValue;
                        SubclassClass MF = new SubclassClass();
                        MF.SubclassName = oListItem["Subclass"].ToString();

                        MF.ClassName = lookup.LookupValue;
                        if (MF.ClassName == "Guru")
                        {
                            MF.ClassAlias = "GG";
                        }
                        else if (MF.ClassName == "Lara")
                        {
                            MF.ClassAlias = "LL";
                        }
                        else
                        {
                            MF.ClassAlias = "CH";
                        }

                        lstMF.Add(MF);
                        lstSubclass.Items.Add(MF.SubclassName);

                    }

                }
            }
            catch (Exception exec)
            {
                logger.WriteEventLog("Please contact admin : " + exec.Message);
                throw exec;
            }
        }

        /// <summary>
        /// Click event of export button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        private void btnExport_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Export will take some time to complete. Press OK to continue.");
            lblMessage.Visible = true;
            lblMessage.Text = "Export started, please wait…";
            List<string> Class = GetSelectedClass();
            try
            {
                if (lstSubclass.SelectedIndex >= 0)
                {
                    writeXml = new WriteXML();
                    repData = new RepositoryData();
                    logger = new Logger();
                    ListItemCollection listItemColl = null;
                    string siteUrl = ConfigurationManager.AppSettings["SiteUrl"];
                    ClientContext clientContext = new ClientContext(siteUrl);
                    List<SubclassClass> listSubclass = GetSelectedSubclass();
                    var selectedSubclass = (from A in listSubclass
                                         select A.SubclassName).ToList();
                    listItemColl = repData.GetDocMetaData(clientContext, "Repository", "Custom Document Set", string.Join(",", selectedSubclass.ToArray()));
                    writeXml.CreateXml(listItemColl, dtRevDate.Value, cbDocType.Text, listSubclass, clientContext);
                    MessageBox.Show("Export completed successfully.");
                    lblMessage.Visible = false;
                }
                else
                {
                    MessageBox.Show("Please select Subclass number.");
                }
            }
            catch (Exception exec)
            {
                logger.WriteEventLog("Please contact admin : " + exec.Message);
                throw exec;
            }
        }

        /// <summary>
        /// Get selected Subclass in Class
        /// </summary>
        /// <returns></returns>
        private List<string> GetSelectedClass()
        {
            List<string> listResult = new List<string>();
            for (int i = 0; i < lstSubclass.SelectedIndices.Count; i++)
            {
                string strSelectedSubclass = lstSubclass.Items[lstSubclass.SelectedIndices[i]].ToString();

                var Subclass = (from LI in lstMF
                             where LI.SubclassName == strSelectedSubclass
                             select LI).ToList();
                foreach (SubclassClass MF in Subclass)
                {
                    listResult.Add(MF.ClassName);
                }
            }
            return listResult.Distinct().ToList();
        }

        /// <summary>
        /// Get selected Subclass in form
        /// </summary>
        /// <returns>list of Subclasss</returns>
        public List<SubclassClass> GetSelectedSubclass()
        {
            List<SubclassClass> listResult = new List<SubclassClass>();
            for (int i = 0; i < lstSubclass.SelectedIndices.Count; i++)
            {
                string strSelectedSubclass = lstSubclass.Items[lstSubclass.SelectedIndices[i]].ToString();

                var Subclass = (from LI in lstMF
                             where LI.SubclassName == strSelectedSubclass
                             select LI).ToList();
                foreach (SubclassClass MF in Subclass)
                {
                    listResult.Add(MF);
                }

            }
            return listResult;
        }



    }
}
