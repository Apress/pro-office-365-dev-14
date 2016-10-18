using Microsoft.Office365.Exchange;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SalesLeadMVC.Models
{
    public class CheckAppointmentsViewModel
    {
        public CheckAppointmentsViewModel()
        {
            appointments = new List<Event>();
        }

        public List<Event> appointments { get; set; }
    }
}