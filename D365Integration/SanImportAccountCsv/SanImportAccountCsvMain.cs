using AuthenticationUtility;
using Microsoft.OData.Client;
using ODataUtility.Microsoft.Dynamics.DataEntities;
using SanCommon.Common;
using SanCommon.Const;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace SanImportAccountCsv
{
    class SanImportAccountCsvMain
    {
        static void Main(string[] args)
        {
            System.Threading.Mutex mtx = null;

            //connect to D365
            var context = ImportCommon.Connect();
            
            //read config
            string strFilePath = ImportCommon.GetAppSetting(BusinessConst.CSV_FILE_PATH);
            string strSuccessPath = ImportCommon.GetAppSetting(BusinessConst.SUC_FILE_PATH);
            string strErrorPath = ImportCommon.GetAppSetting(BusinessConst.ERR_FILE_PATH);
            string strLogPath = ImportCommon.GetAppSetting(BusinessConst.LOG_FILE_PATH);

            //log file
            StringBuilder logMsg = new System.Text.StringBuilder();
            string strLogName = ImportCommon.GetLogFileName(BusinessConst.REDAC_NAME_INVOICE);
            string strfileName = "";
            int iFileCnt = 0;

            Console.WriteLine("logName : {0}", strLogName);
            Console.WriteLine("FilePath :  {0}", strFilePath);

            //mail set
            Console.WriteLine("mail setting :::");
            EmailConf emailConf = new EmailConf();

            try
            {
                string mutextName = "SanImportAccountCsv";

                mtx = new System.Threading.Mutex(false, mutextName);

                mtx.WaitOne();

                //open file
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, BusinessConst.LOG_ST);
                IEnumerable<System.IO.FileInfo> targetCsvfiles = ImportCommon.GetCsvFile(strFilePath);

                Console.WriteLine("start loop  ");

                //import data
                foreach (System.IO.FileInfo targetCsvfile in targetCsvfiles)
                {
                    iFileCnt++;
                    strfileName = targetCsvfile.Name;
                    ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, "File name:" + targetCsvfile.Name);
                    Console.WriteLine("strfileName : {0}", strfileName);

                    //write import data
                    List<string> accRecords = ImportCommon.WriteCsvFile(targetCsvfile);

                    var wkLine = ImportCommon.TrimDoubleQuotationMarks(accRecords[1].Replace("\",\"", "\t")).Split('\t');
                    Console.WriteLine("create customer : ");

                    string comId = wkLine[25];
                    if (wkLine[27].Equals("Tenant"))
                    {
                        comId = comId.Replace(",","");
                    }

                    if (SanODataQuerys.CheckByCustomId(context, comId))
                    {
                        SanODataChangesets.CreateCustomAdv(wkLine, context);
                    }
                    else
                    {
                        Console.WriteLine("customer aleady exist: ");
                        SanODataChangesets.AddCustomAddress(wkLine, context);
                    }

                    Console.WriteLine("create journal : ");
                    SanODataChangesets.CreateGeneralJournalAdv(accRecords, context, logMsg);
                    Console.WriteLine("create done !!");
                    
                    //move file
                    ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, "Done:" + strfileName);
                    ImportCommon.MoveCsvFile(strFilePath + "/" + strfileName, strSuccessPath + "/" + strfileName);
                }


                ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, "File count:" + iFileCnt);
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, BusinessConst.LOG_EN);

            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("err somthing  : {0}", e.InnerException);
                //move file
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_ERROR, "Detail1:" + e.Message);
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_ERROR, "Detail2:" + e.InnerException);
                ImportCommon.MoveCsvFile(strFilePath + "/" + strfileName, strErrorPath + "/" + strfileName);

                SendEmail email = new SendEmail(emailConf);
                email.SendException(e, strLogName);

                if (mtx != null)
                {
                    mtx.ReleaseMutex();
                    mtx = null;
                }
            }
            finally
            {
                Console.WriteLine("check log : {0} ", logMsg.ToString());
                if (iFileCnt > 0)
                {
                    ImportCommon.CreatelogFile(strLogPath + "/" + strLogName, logMsg.ToString());
                }
                Console.WriteLine("end  ");

                if (mtx != null)
                {
                    mtx.ReleaseMutex();
                }
            }
        }
    }
}
