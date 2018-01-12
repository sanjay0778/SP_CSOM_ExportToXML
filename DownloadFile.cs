using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportToXML
{
    public class DownloadFile
    {
        Logger logger = new Logger();
        public void FileRef(ListItem item, ClientContext clientContext, List<string> selectedClass, List<string> selectedMode)
        {
            try
            {
                string strFileRelativeUrl = string.Empty;
                string tempLocation = CreateDirectory(selectedClass, selectedMode);
                if (item != null)
                {

                    strFileRelativeUrl = item["FileRef"].ToString();
                    string fileName = strFileRelativeUrl.Substring(strFileRelativeUrl.LastIndexOf('/') + 1);
                    if (!string.IsNullOrEmpty(strFileRelativeUrl))
                    {
                        // string fileName = System.IO.Path.GetFileName(mf.FileRef);
                        clientContext.ExecuteQuery();
                        //download

                        FileInformation fromFileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, strFileRelativeUrl);
                        clientContext.ExecuteQuery();

                        Stream fs = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, strFileRelativeUrl).Stream;
                        clientContext.ExecuteQuery();
                        byte[] binary = ReadFully(fs);
                        FileStream stream = new FileStream(tempLocation + "\\" + fileName, System.IO.FileMode.Create);
                        BinaryWriter writer = new BinaryWriter(stream);
                        writer.Write(binary);
                        writer.Close();
                    }

                }
            }
            catch (Exception exec)
            {
                logger.WriteEventLog("Please contact admin : " + exec.Message);
                throw exec;
            }

        }

        public static byte[] ReadFully(Stream input)
        {

            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

        public string CreateDirectory(List<string> selectedClass, List<string> selectedMode)
        {
            string fileDirectory = ConfigurationManager.AppSettings["XmlFileLocation"];
            if (!Directory.Exists(fileDirectory))
            {
                Directory.CreateDirectory(fileDirectory);
            }
            //[Mode]_BS (ex. xx_BS).
            string strAbbreviation = string.Join("", selectedClass.ToArray());
            string strMode = string.Join("", selectedMode.ToArray()) + "_BS";
            string downloadLocation = fileDirectory + strAbbreviation + strMode;
            if (!Directory.Exists(downloadLocation))
            {
                Directory.CreateDirectory(downloadLocation);
            }
            return downloadLocation;
        }
    }
}
