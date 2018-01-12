using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportToXML
{
    class Logger
    {
        public void WriteEventLog(string errorMessage)
        {
            StringBuilder errorBuilder = new StringBuilder();
            try
            {

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    string fileDateFormat = "yyyyMMdd";
                    string fileDirectory = ConfigurationManager.AppSettings["XmlFileLocation"];
                    string errorLogFilePath = fileDirectory + "ErrorLog";

                    if (!Directory.Exists(errorLogFilePath))
                    {
                        Directory.CreateDirectory(errorLogFilePath);
                    }

                    errorLogFilePath = errorLogFilePath + @"\ErrorLogFile" + DateTime.Now.ToString(fileDateFormat) + ".log";
                    errorBuilder.Append("-------------------------------------------------------------------------------------------------\n");
                    errorBuilder.AppendLine(System.DateTime.Now.ToString());
                    errorBuilder.AppendLine("Error Message :" + errorMessage);
                    errorBuilder.AppendLine("-------------------------------------------------------------------------------------------------");

                    //Writing to Text File.
                    TextWriter oTW = System.IO.File.AppendText(errorLogFilePath);
                    oTW.WriteLine(errorBuilder.ToString());
                    oTW.Close();
                }

            }
            catch (Exception exec)
            {
                throw exec;
            }
        }

        public void LogMissingTitles(string errorMessage, List<string> selectedClass, List<string> selectedSubClass)
        {
            StringBuilder errorBuilder = new StringBuilder();
            try
            {

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    string fileDirectory = ConfigurationManager.AppSettings["XmlFileLocation"];
                    string errorLogFilePath = fileDirectory + "ErrorLog";

                    if (!Directory.Exists(errorLogFilePath))
                    {
                        Directory.CreateDirectory(errorLogFilePath);
                    }
                    string strAbbreviation = string.Join("", selectedClass.ToArray());
                    string strSubClass = string.Join("", selectedSubClass.ToArray());

                    errorLogFilePath = errorLogFilePath + @"\ErrorLog" + strAbbreviation + strSubClass + ".log";
                    errorBuilder.Append("-----------------------------------------------------------------------------------------------------\n");
                    errorBuilder.AppendLine(System.DateTime.Now.ToString());
                    errorBuilder.AppendLine("Error Message :" + errorMessage);
                    errorBuilder.AppendLine("----------------------------------------------------------------------------------------------------");

                    //Writing to Text File.
                    TextWriter oTW = System.IO.File.AppendText(errorLogFilePath);
                    oTW.WriteLine(errorBuilder.ToString());
                    oTW.Close();
                }

            }
            catch (Exception exec)
            {
                throw exec;
            }
        }
    }
}
