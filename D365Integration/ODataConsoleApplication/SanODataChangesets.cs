using Microsoft.OData.Client;
using ODataUtility.Microsoft.Dynamics.DataEntities;
using SanCommon;
using SanCommon.Common;
using SanCommon.Const;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace SanImportCsv
{
    class SanODataChangesets
    {
        public static void CreateGeneralJournal(List<string> accRecords, Resources context, StringBuilder logMsg)
        {
            for (int i = 1; i < accRecords.Count; i++)
            {
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, "Line" + i + ":" + ImportCommon.TrimDoubleQuotationMarks(accRecords[i].Replace("\",\"", "\t")));
                var wkLine = ImportCommon.TrimDoubleQuotationMarks(accRecords[i].Replace("\",\"", "\t")).Split('\t');

                string str1 = wkLine[0];                //Operation
                string str2 = wkLine[1];                //Company ID
                string str3 = wkLine[2];                //Company Name
                string str4 = wkLine[3];                //Street
                string str5 = wkLine[4];                //Street2
                string str6 = wkLine[5];                //Street3

                //Import to AX
                DataServiceCollection<GeneralLedgerCustInvoiceJournalHeader> generalJournalCollection = new DataServiceCollection<GeneralLedgerCustInvoiceJournalHeader>(context);
                DataServiceCollection<GeneralLedgerCustInvoiceJournalLine> generalJournalLineCollection = new DataServiceCollection<GeneralLedgerCustInvoiceJournalLine>(context);

                //create General Journal
                GeneralLedgerCustInvoiceJournalHeader generalJournal = new GeneralLedgerCustInvoiceJournalHeader();
                generalJournalCollection.Add(generalJournal);
                generalJournal.JournalName = "test jn2";
                generalJournal.JournalBatchNumber = "11223";
                generalJournal.Description = "test1";
                generalJournal.DataAreaId = "USMF";

                GeneralLedgerCustInvoiceJournalLine generalJournalLine = new GeneralLedgerCustInvoiceJournalLine();
                GeneralLedgerCustInvoiceJournalLine generalJournalLine2 = new GeneralLedgerCustInvoiceJournalLine();

                generalJournalLineCollection.Add(generalJournalLine);
                generalJournalLineCollection.Add(generalJournalLine2);

                generalJournalLine.InvoiceDate = new System.DateTime();
                generalJournalLine.AccountType = LedgerJournalACType.Cust;
                generalJournalLine.CustomerAccountDisplayValue = "test0001";
                generalJournalLine.JournalBatchNumber = "11223";
                generalJournalLine.Company = "usmf";
                generalJournalLine.DebitAmount = 100;
                generalJournalLine.OffsetAccountType = LedgerJournalACType.Ledger;
                generalJournalLine.OffsetAccountDisplayValue = "----";
                generalJournalLine.MethodOfPayment = "CHECK";
                generalJournalLine.OffsetCompany = "usmf";
                generalJournalLine.InvoiceId = "test2";
                generalJournalLine.Currency = "USD";
                generalJournalLine.DataAreaId = "USMF";
                generalJournalLine.ApprovedBy = "test";
                generalJournalLine.Approved = NoYes.Yes;

                Console.WriteLine(string.Format("LedgerJournalACType.Vend {0} - !", (int)(LedgerJournalACType.Vend)));
                generalJournalLine2.InvoiceDate = new System.DateTime();
                generalJournalLine2.AccountType = LedgerJournalACType.Vend;
                generalJournalLine2.CustomerAccountDisplayValue = "----";
                generalJournalLine2.JournalBatchNumber = "11223";
                generalJournalLine2.Company = "usmf";
                generalJournalLine2.CreditAmount = 100;
                generalJournalLine2.OffsetAccountType = LedgerJournalACType.Ledger;
                generalJournalLine2.OffsetAccountDisplayValue = "----";
                generalJournalLine2.OffsetCompany = "usmf";
                //generalJournalLine2.InvoiceId = "afeafe";
                generalJournalLine2.Currency = "USD";
                generalJournalLine2.DataAreaId = "USMF";

                context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset); // Batch with Single Changeset ensure the saved changed runs in all-or-nothing mode.
                Console.WriteLine(string.Format("Invoice {0} - Saved !", wkLine[1]));
            }


        }

        public static void CreateGeneralJournalAdv(List<string> accRecords, Resources context, StringBuilder logMsg)
        {
            for (int i = 1; i < accRecords.Count; i++)
            {
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, "Line" + i + ":" + ImportCommon.TrimDoubleQuotationMarks(accRecords[i].Replace("\",\"", "\t")));
                var wkLine = ImportCommon.TrimDoubleQuotationMarks(accRecords[i].Replace("\",\"", "\t")).Split('\t');

                string str1 = wkLine[0];                //Order No
                string str2 = wkLine[1];                //Advanced Payment ID
                string str3 = wkLine[2];                //Submit Date
                string str4 = wkLine[3];                //Payment Date
                string str5 = wkLine[4];                //Payment Time
                string str6 = wkLine[5];                //Total Amount
                string str7 = wkLine[6];                //Redac Division
                string str8 = wkLine[7];                //Parent Redac Division
                string str9 = wkLine[8];                 //Tenant Name
                string str10 = wkLine[9];                //Company Name
                string str11 = wkLine[10];                //Street
                string str12 = wkLine[11];                //Street2
                string str13 = wkLine[12];                //Street3
                string str14 = wkLine[13];                //City
                string str15 = wkLine[14];                //US State
                string str16 = wkLine[15];                //ZIP/Postal code
                string str17 = wkLine[16];                //Country
                string str18 = wkLine[17];                //Main Phone
                string str19 = wkLine[18];                //Fax
                string str20 = wkLine[19];                //Email
                string str21 = wkLine[20];                //Purpose
                string str22 = wkLine[21];                //Amount
                string str23 = wkLine[22];                //Advance/Expense
                string str24 = wkLine[23];                //Payable to
                string str25 = wkLine[24];                //Payment Method
                string str26 = wkLine[25];                //Company ID
                string str27 = wkLine[26];                //Type
                string str28 = wkLine[27];                //CustomerGroup

                string strvendNum = "";
                string strInvoice = str2 + "-" + i;

                Console.WriteLine("create vendor : ");
                strvendNum = SanODataQuerys.CheckByVendorId(context, str2, str24);
                if (strvendNum.Length == 0)
                {
                    strvendNum = CreateVendor(wkLine, context, i);
                }
                else
                {
                    Console.WriteLine("vendor aleady exist: ");
                }

                //Import to AX
                DataServiceCollection<LedgerJournalHeader> generalJournalCollection = new DataServiceCollection<LedgerJournalHeader>(context);

                //create General Journal header
                LedgerJournalHeader generalJournal = new LedgerJournalHeader();
                generalJournalCollection.Add(generalJournal);
                generalJournal.JournalName = ConfigurationManager.AppSettings["JournalName"];

                //Description [Advanced Payment ID]+[Line#] / [Submit Date] / [Amount] / [Purpose]
                string strDescription = strInvoice + "/" + str3.Replace("/","-") + "/" + decimal.Parse(str22).ToString("0.00") + "/" + str21;
                generalJournal.Description = strDescription;
                generalJournal.DataAreaId = ConfigurationManager.AppSettings["RedacCompany"];

                DataServiceResponse res1 = null;
                res1 = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);

                string strJournalbatch = "";
                foreach (ChangeOperationResponse change in res1)
                {
                    // Get the descriptor for the entity.
                    EntityDescriptor descriptor = change.Descriptor as EntityDescriptor;

                    if (descriptor != null)
                    {
                        LedgerJournalHeader LedgerJournalres = descriptor.Entity as LedgerJournalHeader;
                        if (LedgerJournalres != null)
                        {
                            strJournalbatch = LedgerJournalres.JournalBatchNumber;
                            Console.WriteLine("New JournalBatchNumber {0}.", strJournalbatch);
                        }
                    }
                }
                
                //create General Journal line
                DataServiceCollection<LedgerJournalLine> generalJournalLineCollection = new DataServiceCollection<LedgerJournalLine>(context);
                LedgerJournalLine generalJournalLine = new LedgerJournalLine();
                LedgerJournalLine generalJournalLine2 = new LedgerJournalLine();

                generalJournalLineCollection.Add(generalJournalLine);
                generalJournalLineCollection.Add(generalJournalLine2);

                generalJournalLine.TransDate = Convert.ToDateTime(str3);
                generalJournalLine.AccountType = LedgerJournalACType.Cust;
                generalJournalLine.AccountDisplayValue = ImportCommon.Get20CustomerNum(str26).Replace(",", "").Replace("-", "\\-");
                generalJournalLine.Text = str21;
                generalJournalLine.JournalBatchNumber = strJournalbatch;
                generalJournalLine.Company = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine.DebitAmount = Convert.ToDecimal(str22);
                generalJournalLine.OffsetAccountType = LedgerJournalACType.Ledger;
                generalJournalLine.OffsetAccountDisplayValue = "----";
                generalJournalLine.OffsetCompany = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine.Invoice = strInvoice;
                generalJournalLine.CurrencyCode = ConfigurationManager.AppSettings["Currency"];
                generalJournalLine.DataAreaId = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine.PostingProfile = ConfigurationManager.AppSettings["AdvancedPostingCustomer"];

                Console.WriteLine(string.Format("LedgerJournalACType.Vend {0} - !", (int)(LedgerJournalACType.Vend)));
                generalJournalLine2.AccountType = LedgerJournalACType.Vend;
                generalJournalLine2.AccountDisplayValue = strvendNum.Replace("-", "\\-");
                generalJournalLine2.JournalBatchNumber = strJournalbatch;
                generalJournalLine2.Company = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine2.CreditAmount = Convert.ToDecimal(str22);
                generalJournalLine2.OffsetAccountType = LedgerJournalACType.Ledger;
                generalJournalLine2.OffsetAccountDisplayValue = "----";
                generalJournalLine2.PaymentMethod = ConfigurationManager.AppSettings[str25];
                generalJournalLine2.DueDate = Convert.ToDateTime(str4);
                generalJournalLine2.OffsetCompany = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine2.Invoice = strInvoice;
                generalJournalLine2.CurrencyCode = ConfigurationManager.AppSettings["Currency"];
                generalJournalLine2.DataAreaId = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine2.PostingProfile = ConfigurationManager.AppSettings["AdvancedPostingVendor"];

                context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset); // Batch with Single Changeset ensure the saved changed runs in all-or-nothing mode.
                Console.WriteLine(string.Format("Invoice {0} - Saved !", strJournalbatch));

                //SimpleCRUDExamples.GjUpd(context, strJournalbatch);
                //create General Journal line
                /*
                DataServiceCollection<GeneralLedgerCustInvoiceJournalLine> generalInvoiceCollection = new DataServiceCollection<GeneralLedgerCustInvoiceJournalLine>(context);
                GeneralLedgerCustInvoiceJournalLine generalInvoiceLine = new GeneralLedgerCustInvoiceJournalLine();
                generalInvoiceCollection.Add(generalInvoiceLine);

                Console.WriteLine(string.Format("approve Line start - !"));
                generalInvoiceLine.JournalBatchNumber = strJournalbatch;
                generalInvoiceLine.ApprovedBy = "000728";
                generalInvoiceLine.Approved = NoYes.Yes;

                context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset); // Batch with Single Changeset ensure the saved changed runs in all-or-nothing mode.
                Console.WriteLine(string.Format("VendInvoice {0} - Saved !", strJournalbatch));
                */

            }





        }

        public static void CreateGeneralJournalInv(List<string> accRecords, Resources context, StringBuilder logMsg)
        {
            for (int i = 1; i < accRecords.Count; i++)
            {
                ImportCommon.Writelog(logMsg, BusinessConst.LOG_INFO, "Line" + i + ":" + ImportCommon.TrimDoubleQuotationMarks(accRecords[i].Replace("\",\"", "\t")));
                var wkLine = ImportCommon.TrimDoubleQuotationMarks(accRecords[i].Replace("\",\"", "\t")).Split('\t');

                string str1 = wkLine[0];                //Order No
                string str2 = wkLine[1];                //Invoice No
                string str3 = wkLine[2];                //Invoice Date
                string str4 = wkLine[3];                //Total Amount
                string str5 = wkLine[4];                //Attention Salutation
                string str6 = wkLine[5];                //Attention
                string str7 = wkLine[6];                //Redac Attention
                string str8 = wkLine[7];                //Redac Division
                string str9 = wkLine[8];                 //Invoice To
                string str10 = wkLine[9];                //Invoice To Onwer/Other Name
                string str11 = wkLine[10];                //Street
                string str12 = wkLine[11];                //Street2
                string str13 = wkLine[12];                //Street3
                string str14 = wkLine[13];                //City
                string str15 = wkLine[14];                //US State
                string str16 = wkLine[15];                //ZIP/Postal code
                string str17 = wkLine[16];                //Country
                string str18 = wkLine[17];                //Amount
                string str19 = wkLine[18];                //Description
                string str20 = wkLine[19];                //Paid to
                string str21 = wkLine[20];                //Remarks
                string str22 = wkLine[21];                //Div code
                string str23 = wkLine[22];                //Business
                string str24 = wkLine[23];                //State
                string str25 = wkLine[24];                //Local
                string str26 = wkLine[25];                //PL
                string str27 = wkLine[26];                //Main Phone
                string str28 = wkLine[27];                //Fax
                string str29 = wkLine[28];                //Email
                string str30 = wkLine[29];                //Company ID
                string str31 = wkLine[30];                //Type
                string str32 = wkLine[31];                //CustomerGroup

                //set OffsetAccountDisplayValue
                string strOffsetAccount = "";
                string strInvoice = str2 + "-" + i;

                if (str19.Equals("Relocation Service Fee") || str19.Equals("Owner's Pay"))
                {
                    //strOffsetAccount = "40140-IR-Los Angeles-CA-Irvine";
                    strOffsetAccount = ConfigurationManager.AppSettings["InvoiceLedger"] + "-" + str22 + "-" + str26 + "-" + str24 + "-" + str25;
                }
                else
                {
                    strOffsetAccount = ConfigurationManager.AppSettings["InvoiceLedgerOther"] + "-" + str22 + "-" + str26 + "-" + str24 + "-" + str25;
                    //account-Div code-PL-State-Local
                    //strOffsetAccount = "40140-IR-Los Angeles-CA-Irvine";
                }
                //create General Journal Header
                DataServiceCollection<LedgerJournalHeader> generalJournalCollection = new DataServiceCollection<LedgerJournalHeader>(context);
                LedgerJournalHeader generalJournal = new LedgerJournalHeader();
                generalJournalCollection.Add(generalJournal);
                generalJournal.JournalName = ConfigurationManager.AppSettings["JournalName"];

                //Description [Invoice ID]+[Line#] / [Submit Date] / [Amount] / [Purpose]
                string strDescription = strInvoice + "/" + str3.Replace("/", "-") + "/" + decimal.Parse(str18).ToString("0.00") + "/" + str19;
                generalJournal.Description = strDescription;
                generalJournal.DataAreaId = ConfigurationManager.AppSettings["RedacCompany"];

                DataServiceResponse res1 = null;
                res1 = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);

                string strJournalbatch = "";
                foreach (ChangeOperationResponse change in res1)
                {
                    // Get the descriptor for the entity.
                    EntityDescriptor descriptor = change.Descriptor as EntityDescriptor;

                    if (descriptor != null)
                    {
                        LedgerJournalHeader LedgerJournalres = descriptor.Entity as LedgerJournalHeader;
                        if (LedgerJournalres != null)
                        {
                            strJournalbatch = LedgerJournalres.JournalBatchNumber;
                            Console.WriteLine("New JournalBatchNumber {0}.", strJournalbatch);
                        }
                    }
                }

                //create General Journal Line
                DataServiceCollection<LedgerJournalLine> generalJournalLineCollection = new DataServiceCollection<LedgerJournalLine>(context);
                LedgerJournalLine generalJournalLine = new LedgerJournalLine();
                generalJournalLineCollection.Add(generalJournalLine);

                generalJournalLine.TransDate = Convert.ToDateTime(str3);
                generalJournalLine.AccountType = LedgerJournalACType.Cust;
                if (str32.Equals("Tenant"))
                {
                    generalJournalLine.AccountDisplayValue = ImportCommon.Get20CustomerNum(str30).Replace(",","").Replace("-", "\\-");
                }
                else
                {
                    generalJournalLine.AccountDisplayValue = ImportCommon.Get20CustomerNum(str30).Replace("-", "\\-");
                }

                generalJournalLine.Text = str19;
                generalJournalLine.JournalBatchNumber = strJournalbatch;
                generalJournalLine.Company = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine.DebitAmount = Convert.ToDecimal(str18);
                generalJournalLine.OffsetAccountType = LedgerJournalACType.Ledger;
                generalJournalLine.OffsetAccountDisplayValue = strOffsetAccount;
                generalJournalLine.OffsetCompany = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine.Invoice = strInvoice;
                generalJournalLine.CurrencyCode = ConfigurationManager.AppSettings["Currency"];
                generalJournalLine.DataAreaId = ConfigurationManager.AppSettings["RedacCompany"];
                generalJournalLine.PostingProfile = ConfigurationManager.AppSettings["InvoicePostingcustomer"];

                context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset); // Batch with Single Changeset ensure the saved changed runs in all-or-nothing mode.
                Console.WriteLine(string.Format("Invoice {0} - Saved !", str3));
            }


        }

        public static void CreateCustomAdv(String[] wkLine, Resources context)
        {

            string str1 = wkLine[0];                //Order No
            string str2 = wkLine[1];                //Advanced Payment ID
            string str3 = wkLine[2];                //Submit Date
            string str4 = wkLine[3];                //Payment Date
            string str5 = wkLine[4];                //Payment Time
            string str6 = wkLine[5];                //Total Amount
            string str7 = wkLine[6];                //Redac Division
            string str8 = wkLine[7];                //Parent Redac Division
            string str9 = wkLine[8];                 //Tenant Name
            string str10 = wkLine[9];                //Company Name
            string str11 = wkLine[10];                //Street
            string str12 = wkLine[11];                //Street2
            string str13 = wkLine[12];                //Street3
            string str14 = wkLine[13];                //City
            string str15 = wkLine[14];                //US State
            string str16 = wkLine[15];                //ZIP/Postal code
            string str17 = wkLine[16];                //Country
            string str18 = wkLine[17];                //Main Phone
            string str19 = wkLine[18];                //Fax
            string str20 = wkLine[19];                //Email
            string str21 = wkLine[20];                //Purpose
            string str22 = wkLine[21];                //Amount
            string str23 = wkLine[22];                //Advance/Expense
            string str24 = wkLine[23];                //Payable to
            string str25 = wkLine[24];                //Payment Method
            string str26 = wkLine[25];                //Company ID
            string str27 = wkLine[26];                //Type
            string str28 = wkLine[27];                //CustomerGroup

            Customer myCustomer = new Customer();
            DataServiceCollection<Customer> customersCollection = new DataServiceCollection<Customer>(context);
            //DataServiceCollection<CustomerPostalAddress> customersCollection2 = new DataServiceCollection<CustomerPostalAddress>(context);
            customersCollection.Add(myCustomer);

            if (str28.Equals("Company"))
            {
                myCustomer.CustomerAccount = str26;
                myCustomer.CustomerGroupId = ConfigurationManager.AppSettings["CustomerGroupIdCom"];
                myCustomer.PartyType = ConfigurationManager.AppSettings["PartyTypeOrg"];
                myCustomer.Name = str10;
            }
            else
            {
                myCustomer.CustomerGroupId = ConfigurationManager.AppSettings["CustomerGroupIdTen"];
                myCustomer.PartyType = ConfigurationManager.AppSettings["PartyTypePsn"];

                string[] stArrayName = str10.Split(',');
                myCustomer.CustomerAccount = str26.Replace(",", "");
                myCustomer.PersonPhoneticFirstName = stArrayName[0];
                myCustomer.PersonPhoneticLastName = stArrayName[1];
                myCustomer.Name = str10.Replace(",","");
            }

            myCustomer.AddressLocationRoles = "Business";
            myCustomer.SalesCurrencyCode = ConfigurationManager.AppSettings["Currency"];
            myCustomer.AddressDescription = str2;
            myCustomer.AddressStreet = str11 + " " + str12 + " " + str13;
            myCustomer.AddressCity = str14;
            myCustomer.AddressState = ConfigurationManager.AppSettings[str15];
            myCustomer.AddressZipCode = str16;
            myCustomer.AddressCountryRegionId = str17.Length > 0 ? ConfigurationManager.AppSettings[str17] : "USA";

            if (str18.Length > 0)
            {
                myCustomer.PrimaryContactPhoneDescription = str2;
                myCustomer.PrimaryContactPhone = str18;
            }

            if (str19.Length > 0)
            { 
                myCustomer.PrimaryContactFaxDescription = str2;
                myCustomer.PrimaryContactFax = str19;
            }

            if (str20.Length > 0)
            { 
                myCustomer.PrimaryContactEmailDescription = str2;
                myCustomer.PrimaryContactEmail = str20;
            }

            #region Create customers address
            /*
            CustomerPostalAddress myCustomer2 = new CustomerPostalAddress();
            customersCollection2.Add(myCustomer2);

            myCustomer2.CustomerAccountNumber = "0001";
            myCustomer2.CustomerLegalEntityId = "RRI";
            myCustomer2.AddressLocationRoles = "Business";
            //myCustomer2.AttentionToAddressLine = "test";
            //myCustomer2.IsRoleBusiness = NoYes.Yes;
            //myCustomer2.AddressLocationId = "";
            myCustomer2.AddressCountryRegionId = "USA";
            myCustomer2.AddressDescription = str10 + "\n" + str11 + "\n" + str12 + " " + str13 + " " + str14;
            myCustomer2.AddressCity = str15;
            myCustomer2.IsPrimary = NoYes.No;
            myCustomer2.IsRoleBusiness = NoYes.Yes;
            myCustomer2.IsPostalAddress = NoYes.Yes;
            //myCustomer2.AddressStreet = "test 5555";
            //myCustomer2.AddressCity = "Barrow";
            //myCustomer2.AddressState = "AK";
            */
            #endregion

            DataServiceResponse response = null;
            try
            {
                response = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);
                Console.WriteLine("created ok");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.InnerException);
            }
        }

        public static void AddCustomAddress(String[] wkLine, Resources context)
        {

            string str1 = wkLine[0];                //Order No
            string str2 = wkLine[1];                //Advanced Payment ID
            string str3 = wkLine[2];                //Submit Date
            string str4 = wkLine[3];                //Payment Date
            string str5 = wkLine[4];                //Payment Time
            string str6 = wkLine[5];                //Total Amount
            string str7 = wkLine[6];                //Redac Division
            string str8 = wkLine[7];                //Parent Redac Division
            string str9 = wkLine[8];                 //Tenant Name
            string str10 = wkLine[9];                //Company Name
            string str11 = wkLine[10];                //Street
            string str12 = wkLine[11];                //Street2
            string str13 = wkLine[12];                //Street3
            string str14 = wkLine[13];                //City
            string str15 = wkLine[14];                //US State
            string str16 = wkLine[15];                //ZIP/Postal code
            string str17 = wkLine[16];                //Country
            string str18 = wkLine[17];                //Main Phone
            string str19 = wkLine[18];                //Fax
            string str20 = wkLine[19];                //Email
            string str21 = wkLine[20];                //Purpose
            string str22 = wkLine[21];                //Amount
            string str23 = wkLine[22];                //Advance/Expense
            string str24 = wkLine[23];                //Payable to
            string str25 = wkLine[24];                //Payment Method
            string str26 = wkLine[25];                //Company ID
            string str27 = wkLine[26];                //Type
            string str28 = wkLine[27];                //CustomerGroup

            Console.WriteLine("account id : {0}", str26);
            DataServiceCollection<CustomerPostalAddress> customersCollection = new DataServiceCollection<CustomerPostalAddress>(context);
            CustomerPostalAddress myCustomer = new CustomerPostalAddress();
            customersCollection.Add(myCustomer);

            myCustomer.CustomerAccountNumber = str26.Replace(",","");
            myCustomer.CustomerLegalEntityId = ConfigurationManager.AppSettings["RedacCompany"];
            myCustomer.AddressLocationRoles = "Business";
            myCustomer.AddressDescription = str2;
            myCustomer.AddressStreet = str11 + " " + str12 + " " + str13;
            myCustomer.AddressCity = str14;
            myCustomer.AddressState = ConfigurationManager.AppSettings[str15];
            myCustomer.AddressZipCode = str16;
            myCustomer.AddressCountryRegionId = str17.Length > 0 ? ConfigurationManager.AppSettings[str17] : "USA";
            myCustomer.IsPrimary = NoYes.No;

            DataServiceResponse response = null;
            try
            {
                response = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);
                Console.WriteLine("added ok");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.InnerException);
            }
        }

        public static void CreateCustomInv(String[] wkLine, Resources context)
        {

            string str1 = wkLine[0];                //Order No
            string str2 = wkLine[1];                //Invoice No
            string str3 = wkLine[2];                //Invoice Date
            string str4 = wkLine[3];                //Total Amount
            string str5 = wkLine[4];                //Attention Salutation
            string str6 = wkLine[5];                //Attention
            string str7 = wkLine[6];                //Redac Attention
            string str8 = wkLine[7];                //Redac Division
            string str9 = wkLine[8];                 //Invoice To
            string str10 = wkLine[9];                //Invoice To Onwer/Other Name
            string str11 = wkLine[10];                //Street
            string str12 = wkLine[11];                //Street2
            string str13 = wkLine[12];                //Street3
            string str14 = wkLine[13];                //City
            string str15 = wkLine[14];                //US State
            string str16 = wkLine[15];                //ZIP/Postal code
            string str17 = wkLine[16];                //Country
            string str18 = wkLine[17];                //Amount
            string str19 = wkLine[18];                //Description
            string str20 = wkLine[19];                //Paid to
            string str21 = wkLine[20];                //Remarks
            string str22 = wkLine[21];                //Div code
            string str23 = wkLine[22];                //Business
            string str24 = wkLine[23];                //State
            string str25 = wkLine[24];                //Local
            string str26 = wkLine[25];                //PL
            string str27 = wkLine[26];                //Main Phone
            string str28 = wkLine[27];                //Fax
            string str29 = wkLine[28];                //Email
            string str30 = wkLine[29];                //Company ID
            string str31 = wkLine[30];                //Type
            string str32 = wkLine[31];                //CustomerGroup

            Customer myCustomer = new Customer();
            DataServiceCollection<Customer> customersCollection = new DataServiceCollection<Customer>(context);
            customersCollection.Add(myCustomer);

            if (str32.Equals("Company"))
            {
                myCustomer.CustomerAccount = str30;
                myCustomer.CustomerGroupId = ConfigurationManager.AppSettings["CustomerGroupIdCom"];
                myCustomer.PartyType = ConfigurationManager.AppSettings["PartyTypeOrg"];
                myCustomer.Name = str10;
            }
            else if (str32.Equals("Tenant"))
            {
                myCustomer.CustomerGroupId = ConfigurationManager.AppSettings["CustomerGroupIdTen"];
                myCustomer.PartyType = ConfigurationManager.AppSettings["PartyTypePsn"];

                string[] stArrayName = str10.Split(',');
                myCustomer.CustomerAccount = str30.Replace(",", "");
                myCustomer.PersonPhoneticFirstName = stArrayName[0];
                myCustomer.PersonPhoneticLastName = stArrayName[1];
                myCustomer.Name = str10.Replace(",", "");
            }
            else if (str32.Equals("Owner"))
            {
                myCustomer.CustomerGroupId = ConfigurationManager.AppSettings["CustomerGroupIdOwn"];
                myCustomer.PartyType = ConfigurationManager.AppSettings["PartyTypeOrg"];
                myCustomer.CustomerAccount = str30;
                myCustomer.Name = str10;
            }
            else
            {
                myCustomer.CustomerGroupId = ConfigurationManager.AppSettings["CustomerGroupIdOth"];
                myCustomer.PartyType = ConfigurationManager.AppSettings["PartyTypeOrg"];
                myCustomer.CustomerAccount = str30;
                myCustomer.Name = str10;
            }

            myCustomer.AddressLocationRoles = "Business";
            myCustomer.SalesCurrencyCode = ConfigurationManager.AppSettings["Currency"];
            myCustomer.AddressDescription = str2;
            myCustomer.AddressStreet = str11 + " " + str12 + " " + str13;
            myCustomer.AddressCity = str14;
            myCustomer.AddressState = ConfigurationManager.AppSettings[str15];
            myCustomer.AddressZipCode = str16;
            myCustomer.AddressCountryRegionId = str17.Length > 0 ? ConfigurationManager.AppSettings[str17] : "USA";

            if (str27.Length > 0)
            {
                myCustomer.PrimaryContactPhoneDescription = str2;
                myCustomer.PrimaryContactPhone = str27;
            }

            if (str28.Length > 0)
            {
                myCustomer.PrimaryContactFaxDescription = str2;
                myCustomer.PrimaryContactFax = str28;
            }

            if (str29.Length > 0)
            {
                myCustomer.PrimaryContactEmailDescription = str2;
                myCustomer.PrimaryContactEmail = str29;
            }

            DataServiceResponse response = null;
            try
            {
                response = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);
                Console.WriteLine("created ok");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.InnerException);
            }
        }

        public static void AddCustomInv(String[] wkLine, Resources context)
        {

            string str1 = wkLine[0];                //Order No
            string str2 = wkLine[1];                //Invoice No
            string str3 = wkLine[2];                //Invoice Date
            string str4 = wkLine[3];                //Total Amount
            string str5 = wkLine[4];                //Attention Salutation
            string str6 = wkLine[5];                //Attention
            string str7 = wkLine[6];                //Redac Attention
            string str8 = wkLine[7];                //Redac Division
            string str9 = wkLine[8];                 //Invoice To
            string str10 = wkLine[9];                //Invoice To Onwer/Other Name
            string str11 = wkLine[10];                //Street
            string str12 = wkLine[11];                //Street2
            string str13 = wkLine[12];                //Street3
            string str14 = wkLine[13];                //City
            string str15 = wkLine[14];                //US State
            string str16 = wkLine[15];                //ZIP/Postal code
            string str17 = wkLine[16];                //Country
            string str18 = wkLine[17];                //Amount
            string str19 = wkLine[18];                //Description
            string str20 = wkLine[19];                //Paid to
            string str21 = wkLine[20];                //Remarks
            string str22 = wkLine[21];                //Div code
            string str23 = wkLine[22];                //Business
            string str24 = wkLine[23];                //State
            string str25 = wkLine[24];                //Local
            string str26 = wkLine[25];                //PL
            string str27 = wkLine[26];                //Main Phone
            string str28 = wkLine[27];                //Fax
            string str29 = wkLine[28];                //Email
            string str30 = wkLine[29];                //Company ID
            string str31 = wkLine[30];                //Type
            string str32 = wkLine[31];                //CustomerGroup

            Console.WriteLine("account id : {0}", str30);
            DataServiceCollection<CustomerPostalAddress> customersCollection = new DataServiceCollection<CustomerPostalAddress>(context);
            CustomerPostalAddress myCustomer = new CustomerPostalAddress();
            customersCollection.Add(myCustomer);

            if (str32.Equals("Tenant"))
            {
                myCustomer.CustomerAccountNumber = str30.Replace(",","");
            }
            else
            {
                myCustomer.CustomerAccountNumber = str30;
            }
            myCustomer.CustomerLegalEntityId = ConfigurationManager.AppSettings["RedacCompany"];
            myCustomer.AddressLocationRoles = "Business";
            myCustomer.AddressCountryRegionId = str17.Length > 0 ? ConfigurationManager.AppSettings[str17] : "USA";
            myCustomer.AddressDescription = str2;
            myCustomer.AddressStreet = str11 + " " + str12 + " " + str13;
            myCustomer.AddressCity = str14;
            myCustomer.AddressState = ConfigurationManager.AppSettings[str15];
            myCustomer.AddressZipCode = str16;
            myCustomer.IsPrimary = NoYes.No;

            DataServiceResponse response = null;
            try
            {
                response = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);
                Console.WriteLine("created ok");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.InnerException);
            }
        }

        public static string CreateVendor(String[] wkLine, Resources context, int seq)
        {
            string str1 = wkLine[0];                //Order No
            string str2 = wkLine[1];                //Advanced Payment ID
            string str3 = wkLine[2];                //Submit Date
            string str4 = wkLine[3];                //Payment Date
            string str5 = wkLine[4];                //Payment Time
            string str6 = wkLine[5];                //Total Amount
            string str7 = wkLine[6];                //Redac Division
            string str8 = wkLine[7];                //Parent Redac Division
            string str9 = wkLine[8];                 //Tenant Name
            string str10 = wkLine[9];                //Company Name
            string str11 = wkLine[10];                //Street
            string str12 = wkLine[11];                //Street2
            string str13 = wkLine[12];                //Street3
            string str14 = wkLine[13];                //City
            string str15 = wkLine[14];                //US State
            string str16 = wkLine[15];                //ZIP/Postal code
            string str17 = wkLine[16];                //Country
            string str18 = wkLine[17];                //Main Phone
            string str19 = wkLine[18];                //Fax
            string str20 = wkLine[19];                //Email
            string str21 = wkLine[20];                //Purpose
            string str22 = wkLine[21];                //Amount
            string str23 = wkLine[22];                //Advance/Expense
            string str24 = wkLine[23];                //Payable to
            string str25 = wkLine[24];                //Payment Method

            //string strVendor = str2 + seq.ToString("00");
            string strVendor = str24;
            Vendor myVendor = new Vendor();
            DataServiceCollection<Vendor> vendorsCollection = new DataServiceCollection<Vendor>(context);

            vendorsCollection.Add(myVendor);

            //myVendor.VendorAccountNumber = strVendor;
            myVendor.VendorAccountNumber = str24;
            myVendor.VendorName = str24;
            myVendor.VendorGroupId = ConfigurationManager.AppSettings["VendorGroupId"];
            myVendor.VendorPartyType = ConfigurationManager.AppSettings["VendorPartyType"];

            DataServiceResponse response = null;
            try
            {
                response = context.SaveChanges(SaveChangesOptions.PostOnlySetProperties | SaveChangesOptions.BatchWithSingleChangeset);
                
                Console.WriteLine("created vendor ok");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.InnerException);
            }
            return strVendor;
        }

 
        public static void WriteCsv(List<String[]> wkLine)
        {
            try
            {
                string csvPath = "";
                System.Text.StringBuilder csvFileName = new System.Text.StringBuilder();
                System.Text.StringBuilder csvHeader = new System.Text.StringBuilder();
                System.Text.StringBuilder csvText = new System.Text.StringBuilder();


                String[] mwLiine = new String[20];
                mwLiine = wkLine[0];

                //set filename''';
                csvFileName.Append(BusinessConst.REDAC_NAME_INVOICE);
                csvFileName.Append("_");
                csvFileName.Append(mwLiine[0]);
                csvFileName.Append("_");
                csvFileName.Append(DateTime.Now.ToString("yyyyMMddHHmmss"));
                csvFileName.Append(".csv");

                csvPath = ImportCommon.GetAppSetting(BusinessConst.CSV_FILE_PATH) + csvFileName.ToString();

                //set csv header
                csvHeader.Append(BusinessConst.CSV_DQ + "Order No" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Advanced Payment ID" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Submit Date" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Payment Date" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Payment Time" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Total Amount" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Redac Division" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Parent Redac Division" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Tenant Name" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Company Name" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Street" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Street2" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Street3" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "City" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "US State" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "ZIP/Postal code" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Country" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Main Phone" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Fax" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Email" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Purpose" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Amount" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Advance/Expense" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Payable to" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Payment Method" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Company ID" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "Type" + BusinessConst.CSV_DQ);
                csvHeader.Append(",").Append(BusinessConst.CSV_DQ + "CustomerGroup" + BusinessConst.CSV_DQ + Environment.NewLine);

                //set csv body
                for (int i = 0; i < wkLine.Count; i++)
                {
                    String[] mwLiinefor = new String[20];
                    mwLiinefor = wkLine[i];

                    csvText.Append(BusinessConst.CSV_DQ + mwLiinefor[1] + BusinessConst.CSV_DQ);
                    csvText.Append(",").Append(BusinessConst.CSV_DQ + mwLiinefor[1] + BusinessConst.CSV_DQ);
                    csvText.Append(",").Append(BusinessConst.CSV_DQ + mwLiinefor[1] + BusinessConst.CSV_DQ);
                    csvText.Append(",").Append(BusinessConst.CSV_DQ + mwLiinefor[1] + BusinessConst.CSV_DQ);
                    csvText.Append(",").Append(BusinessConst.CSV_DQ + mwLiinefor[1] + BusinessConst.CSV_DQ);
                    csvText.Append(",").Append(BusinessConst.CSV_DQ + mwLiinefor[1].Replace(BusinessConst.CSV_DQ, "\"\"") + BusinessConst.CSV_DQ);
                    csvText.Append(",").Append(BusinessConst.CSV_DQ + mwLiinefor[1] + BusinessConst.CSV_DQ + Environment.NewLine);
                }

                ImportCommon.CreateCsv(csvPath, csvHeader.ToString(), csvText.ToString());
            }
            catch (System.Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
