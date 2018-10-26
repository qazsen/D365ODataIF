using AuthenticationUtility;
using Microsoft.OData.Client;
using ODataUtility.Microsoft.Dynamics.DataEntities;
using SanCommon.Const;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Net;

namespace SanCommon.Common
{
    public class ImportCommon
    {
        public static Resources Connect()
        {
            String[] appCon = new String[7];

            appCon[0] = GetAppSetting("ServiceUri");
            appCon[1] = GetAppSetting("UserName");
            appCon[2] = GetAppSetting("Password");
            appCon[3] = GetAppSetting("ActiveDirectoryResource");
            appCon[4] = GetAppSetting("ActiveDirectoryTenant");
            appCon[5] = GetAppSetting("ActiveDirectoryClientAppId");
            appCon[6] = GetAppSetting("ActiveDirectoryClientAppSecret");
 

            string ODataEntityPath = appCon[0] + "data";
            Uri oDataUri = new Uri(ODataEntityPath, UriKind.Absolute);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            var context = new Resources(oDataUri);

            context.SendingRequest2 += new EventHandler<SendingRequest2EventArgs>(delegate (object sender, SendingRequest2EventArgs e)
            {
                var authenticationHeader = OAuthHelper.GetAuthenticationHeader2(true, appCon);
                e.RequestMessage.SetHeader(OAuthHelper.OAuthHeader, authenticationHeader);
            });

            return context;

        }

        public static IEnumerable<System.IO.FileInfo> GetCsvFile(string strFilePath)
        {

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(strFilePath);
            IEnumerable<System.IO.FileInfo> files = di.EnumerateFiles("*", System.IO.SearchOption.AllDirectories);

            return files;

        }

        public static string GetAppSetting(string strKey)
        {
            ExeConfigurationFileMap efm = new ExeConfigurationFileMap { ExeConfigFilename = @"..\ComApp.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(efm, ConfigurationUserLevel.None);

            AppSettingsSection appSettings = configuration.AppSettings;
            KeyValueConfigurationElement element1 = appSettings.Settings[strKey];

            return element1.Value;

        }

        public static List<string> WriteCsvFile(System.IO.FileInfo files)
        {
            //write data
            List<string> records = new List<string>();
            using (System.IO.StreamReader r = new System.IO.StreamReader(files.FullName))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    records.Add(line);
                }
            }

            return records;

        }

        public static string TrimDoubleQuotationMarks(string target)
        {
            string field = target.Substring(1, target.Length - 2).Replace("\"\"", "\"");
            return field;
        }

        public static void MoveCsvFile(string source, string target)
        {
            System.IO.File.Move(source, target);
        }
         
        public static void CreatelogFile(string files, string strMsg)
        {
            try
            {
                System.IO.StreamWriter sw = new System.IO.StreamWriter(
                                            files,
                                            false,
                                            Encoding.GetEncoding("utf-8"));
                //write to logfile
                sw.Write(strMsg);
                sw.Close();

            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public static void Writelog(System.Text.StringBuilder strMsg, string info, string addMsg)
        {
            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            strMsg.Append("[").Append(now).Append("]");
            strMsg.Append("[").Append(info).Append("]").Append(addMsg + Environment.NewLine);

            Console.WriteLine("{0} : {1}", now, addMsg);
        }

        public static string GetLogFileName(string entity)
        {
            return BusinessConst.LOG_FILE_NAME + "_" +
                                entity + DateTime.Now.ToString("yyyyMMddHHmmss") + BusinessConst.LOG_FILE_EXTENSION;
        }

        public static string Get20CustomerNum(string target)
        {
            string field = "";
            if (target.Length > 20)
            { 
                field = target.Substring(0, 20);
            }
            else
            {
                field = target;
            }
            return field;
        }

        public static void CreateCsv(string csvPath, string csvHeader, string csvText)
        {
            try
            {
                System.IO.StreamWriter sw = new System.IO.StreamWriter(
                                            csvPath,
                                            false,
                                            System.Text.Encoding.GetEncoding("utf-8"));

                //write to csv
                sw.Write(csvHeader.ToString() + csvText.ToString());
                sw.Close();

            }
            catch (System.Exception e)
            {
                throw new System.Exception(e.Message);
            }
        }

    }
}
