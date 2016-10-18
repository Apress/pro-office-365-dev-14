using Microsoft.Office365.Exchange;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SalesLeadMVC.Models
{
    public class LeadContactViewModel
    {
        public LeadContactViewModel()
        {
            leads = new List<LeadInfo>();
            salesStaff = new List<SalesPerson>();
        }

        public List<LeadInfo> leads { get; set; }
        public List<SalesPerson> salesStaff { get; set; }

        public string selectedLeadID { get; set; }
        public string selectedSalesStaffID { get; set; }
    }

    public class LeadInfo
    {
        public string ID { get; set; }
        public string sku { get; set; }
        public string email { get; set; }
        public string message { get; set; }
        public DateTime dateReceived { get; set; }
        public string productRequest { get; set; }
    }

    public class SalesPerson
    {
        public string ID { get; set; }
        public string name { get; set; }
        public string email { get; set; }
    }

    public class ForwardMessage
    {
        public ForwardMessage()
        {
            ToRecipients = new List<Recipient>();
        }
        public string Comment { get; set; }
        public List<Recipient> ToRecipients { get; set; }
    }

    public class Product
    {
        public string Title { get; set; }
    }
}