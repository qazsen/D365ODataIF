﻿using Microsoft.OData.Client;
using ODataUtility.Microsoft.Dynamics.DataEntities;
using System;
using System.Linq;

namespace SanImportAccountCsv
{
    public static class SanODataQuerys
    {
        public static void ReadLegalEntities(Resources d365)
        {            
            foreach (var legalEntity in d365.LegalEntities.AsEnumerable())
            {
                Console.WriteLine("Name: {0}", legalEntity.Name);
            }
        }

        public static void GetInlineQueryCount(Resources d365)
        {
            var vendorsQuery = d365.Vendors.IncludeTotalCount();
            var vendors = vendorsQuery.Execute() as QueryOperationResponse<Vendor>;

            Console.WriteLine("Total vendors is {0}", vendors.TotalCount);
        }

        public static void GetTopRecords(Resources d365)
        {
            var vendorsQuery = d365.Vendors.AddQueryOption("$top", "10");
            var vendors = vendorsQuery.Execute() as QueryOperationResponse<Vendor>;

            foreach (var vendor in vendors)
            {
                Console.WriteLine("Vendor with ID {0} retrived.", vendor.VendorAccountNumber);
            }
        }

        public static void FilterLinqSyntax(Resources d365)
        {
            var vendors = d365.Vendors.Where(x => x.VendorAccountNumber == "1001");

            foreach (var vendor in vendors)
            {
                Console.WriteLine("Vendor with ID {0} retrived.", vendor.VendorAccountNumber);
            }
        }

        public static void FilterSyntax(Resources d365)
        {
            var vendors = d365.Vendors.AddQueryOption("$filter", "VendorAccountNumber eq '1001'").Execute();

            foreach (var vendor in vendors)
            {
                Console.WriteLine("Vendor with ID {0} retrived.", vendor.VendorAccountNumber);
            }
        }

        public static void SortSyntax(Resources d365)
        {
            var vendors = d365.Vendors.OrderBy(x => x.VendorAccountNumber);

            foreach (var vendor in vendors)
            {
                Console.WriteLine("Vendor with ID {0} retrived.", vendor.VendorAccountNumber);

            }
        }

        public static void FilterByCompany(Resources d365)
        {
            var vendors = d365.Vendors.Where(x => x.DataAreaId == "USMF");

            foreach (var vendor in vendors)
            {
                Console.WriteLine("Vendor with ID {0} retrived.", vendor.VendorAccountNumber);
            }
        }

        public static string CheckByVendorId(Resources d365, string vendorId, string vendorNm)
        {
            string strVendor = "";
            //string wkVendor = "";
            var vendors = d365.Vendors.Where(x => x.VendorName == vendorNm);
            Console.WriteLine("Vendor retrived : {0}", vendors.Count()); 
            //foreach (var vendor in vendors)
            if(vendors.Count() > 0)
            {
                //wkVendor = vendorNm;
                //if (wkVendor.Contains(vendorId))
                //{
                Console.WriteLine("Vendor with ID retrived.");
                    //strVendor = wkVendor;
                strVendor = vendorNm;
                    //break;
                //}
            }
            return strVendor;
        }

        public static bool CheckByCustomId(Resources d365, string coustomId)
        {
            var customers = d365.Customers.Where(x => x.CustomerAccount == coustomId);

            if (customers.Count() > 0)
            {
                Console.WriteLine("Customer with ID retrived.");
                return false;
            }
            else
            {
                return true;
            }

        }


        public static string GetJournalBatchNumber(Resources d365)
        {
            string strJournalBatchNum = "";
            var ledgerJournal = d365.LedgerJournalHeaders.OrderByDescending(x => x.JournalBatchNumber).First();

            strJournalBatchNum = ledgerJournal.JournalBatchNumber;
            Console.WriteLine("JournalBatchNum : {0} retrived.", strJournalBatchNum);

            return strJournalBatchNum;
        }

        public static void ExpandNavigationalProperty(Resources d365Client)
        {            
            var salesOrdersWithLines = d365Client.SalesOrderHeaders.Expand("SalesOrderLine").Where(x => x.SalesOrderNumber == "012518" ).Take(5);

            foreach(var salesOrder in salesOrdersWithLines)
            {
                Console.WriteLine(string.Format("Sales order ID is {0}", salesOrder.SalesOrderNumber));

                foreach( var salesLine in salesOrder.SalesOrderLine)
                {
                    Console.WriteLine(string.Format("Sales order line with description {0} contains item id {1}", salesLine.LineDescription, salesLine.ItemNumber));
                }                
            }
        }

        public static void FilterOnNavigationalProperty(Resources d365Client)
        {
            var salesOrderLines = d365Client.SalesOrderLines.Where(x => x.SalesOrderHeader.SalesOrderStatus == SalesStatus.Invoiced);

            foreach (var salesOrderLine in salesOrderLines)
            {
                Console.WriteLine(salesOrderLine.ItemNumber);
            }          

        }

        public static void FilterLinqTest(Resources d365)
        {
            var vendors = d365.GeneralLedgerCustInvoiceJournalHeaders.AddQueryOption("$top", "10");

            foreach (var vendor in vendors)
            {
                Console.WriteLine("JournalName with ID {0} retrived.", vendor.JournalBatchNumber);
            }
        }
    }
}
