using Microsoft.Office365.Exchange;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SalesLeadMVC.Models
{
    public class LeadAppointmentViewModel
    {
        public LeadAppointmentViewModel()
        {
            appointmentDate = DateTime.Now;
            currentAppointments = new List<Event>();
        }

        public string leadID { get; set; }
        public DateTime appointmentDate { get; set; }
        public string appointmentMessage { get; set; }
        public string appointmentTitle { get; set; }

        public List<Event> currentAppointments { get; set; }
    }
}