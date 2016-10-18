using Microsoft.Office365.Exchange;
using Microsoft.Office365.OAuth;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SalesLeadMVC.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace SalesLeadMVC.Controllers
{
    public class SalesLeadController : Controller
    {
        const string ExchangeResourceId = "https://outlook.office365.com";
        const string ExchangeServiceRoot = "https://outlook.office365.com/ews/odata";
        const string SharePointResourceId = "https://apress365.sharepoint.com";
        const string SharePointServiceRoot = "https://apress365.sharepoint.com/_api";

        private async Task<ExchangeClient> GetExchangeClient()
        {
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(ExchangeResourceId);

            return new ExchangeClient(new Uri(ExchangeServiceRoot), authInfo.GetAccessToken);
        }

        private async Task<Message> GetMessageByID(string messageID)
        {
            //get authorization token
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(ExchangeResourceId);

            //send request through HttpClient
            HttpClient httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri(ExchangeResourceId);
            
            //add authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await authInfo.GetAccessToken());

            //send request
            var response = httpClient.GetAsync("/ews/odata/Me/Inbox/Messages('" + messageID + "')").Result;

            //process response
            if (response.StatusCode == HttpStatusCode.OK)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                var message = JObject.Parse(responseContent).ToObject<Message>();

                return message;
            }

            return null;
        }

        private async Task<bool> ForwardMessage(string messageID, string recipientName, string recipientAddress, string forwardMessage)
        {
            //get authorization token
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(ExchangeResourceId);

            //send request through HttpClient
            HttpClient httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri(ExchangeResourceId);

            //add authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await authInfo.GetAccessToken());

            ForwardMessage forwardContent = new Models.ForwardMessage();
            forwardContent.Comment = forwardMessage;
            forwardContent.ToRecipients.Add(new Recipient() { Address = recipientAddress, Name = recipientName });
            
            StringContent postContent = new StringContent(JsonConvert.SerializeObject(forwardContent));
            postContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            //send request
            var response = httpClient.PostAsync("/ews/odata/Me/Inbox/Messages('" + messageID + "')/Forward", postContent).Result;

            if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.Accepted)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private async Task<List<Product>> GetProducts()
        {
            //get authorization token
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(SharePointResourceId);

            HttpClient httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri(SharePointResourceId);
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await authInfo.GetAccessToken());
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");

            //string listID = "4C13B181-6EA4-42A5-937C-630BCAB28596";
            var listResponse = httpClient.GetAsync("/SalesLeads/_api/web/lists/getbytitle('Products')/items").Result;

            var responseContent = await listResponse.Content.ReadAsStringAsync();

            var json = JObject.Parse(responseContent);

            List<Product> products = json["value"].ToObject<List<Product>>();

            return products;
        }

        private async Task<bool> DeleteMessage(string messageID)
        {
            //get authorization token
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(ExchangeResourceId);

            //send request through HttpClient
            HttpClient httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri(ExchangeResourceId);

            //add authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await authInfo.GetAccessToken());

            //send request
            var response = httpClient.DeleteAsync("/ews/odata/Me/Inbox/Messages('" + messageID + "')").Result;

            if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.NoContent)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //
        // GET: /SalesLead/
        public async Task<ActionResult> Index()
        {
            LeadContactViewModel vm = new LeadContactViewModel();

            var client = await GetExchangeClient();

            var messageResults = await (from i in client.Me.Inbox.Messages
                                        where i.Subject.Contains("[SKU:")
                                        orderby i.DateTimeSent descending
                                        select i).ExecuteAsync();

            foreach (var message in messageResults.CurrentPage)
            {
                LeadInfo newContact = new LeadInfo();

                newContact.ID = message.Id;
                newContact.email = message.From.Address;
                newContact.message = message.BodyPreview;
                newContact.dateReceived = message.DateTimeReceived.Value.DateTime;

                //look up product from SKU
                var beginPos = message.Subject.IndexOf("[", 0) + 1;
                var endPos = message.Subject.IndexOf("]", beginPos);
                string skuLine = message.Subject.Substring(beginPos, endPos - beginPos);

                List<Product> products = await GetProducts();
                var product = products.Where(p => p.Title.Contains(skuLine.Split(':')[1])).SingleOrDefault();

                newContact.productRequest = product.Title;

                vm.leads.Add(newContact);
            }

            //get the sales staff from outlook
            var contactsResults = await (from i in client.Me.Contacts
                                         orderby i.DisplayName
                                         select i).ExecuteAsync();

            foreach (var contact in contactsResults.CurrentPage)
            {
                SalesPerson person = new SalesPerson();

                person.ID = contact.Id;
                person.name = contact.DisplayName;
                person.email = contact.EmailAddress1;

                vm.salesStaff.Add(person);
            }

            return View(vm);
        }

        [HttpPost]
        public async Task<ActionResult> Index(LeadContactViewModel vm)
        {
            var client = await GetExchangeClient();

            var contactResults = await (from i in client.Me.Contacts
                                       select i).ExecuteAsync();

            var contact = contactResults.CurrentPage.Where(c => c.Id == vm.selectedSalesStaffID).SingleOrDefault();

            //USE HTTPCLIENT
            //get authorization token
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(ExchangeResourceId);

            //send request through HttpClient
            HttpClient httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri(ExchangeResourceId);

            //add authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await authInfo.GetAccessToken());

            //prepare POST content for forward request
            ForwardMessage forwardContent = new Models.ForwardMessage();
            forwardContent.Comment = "This lead has been reassigned to you";
            forwardContent.ToRecipients.Add(new Recipient() { Address = contact.EmailAddress1, Name = contact.DisplayName });

            //convert POST content to JSON
            StringContent postContent = new StringContent(JsonConvert.SerializeObject(forwardContent));
            postContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            //send forward request
            var forwardResponse = httpClient.PostAsync("/ews/odata/Me/Inbox/Messages('" + vm.selectedLeadID + "')/Forward", postContent).Result;

            //delete message to remove from INBOX/Lead List
            //send delete request
            var deleteResponse = httpClient.DeleteAsync("/ews/odata/Me/Inbox/Messages('" + vm.selectedLeadID + "')").Result;
            
            //refresh leads
            vm.leads.Clear();
            var messageResults = await (from i in client.Me.Inbox.Messages
                                        where i.Subject.Contains("[SKU:")
                                        orderby i.DateTimeSent descending
                                        select i).ExecuteAsync();

            foreach (var message in messageResults.CurrentPage)
            {
                LeadInfo newContact = new LeadInfo();

                newContact.ID = message.Id;
                newContact.email = message.From.Address;
                newContact.message = message.BodyPreview;
                newContact.dateReceived = message.DateTimeReceived.Value.DateTime;

                vm.leads.Add(newContact);
            }

            return View(vm);
        }

        [HttpGet]
        public async Task<ActionResult> CheckAppointments(DateTime appointmentDate)
        {
            CheckAppointmentsViewModel vm = new CheckAppointmentsViewModel();

            var client = await GetExchangeClient();
            var appointmentResults = await (from i in client.Me.Calendar.Events
                                        where i.Start >= new DateTimeOffset(appointmentDate)
                                        select i).ExecuteAsync();

            foreach (Event appointment in appointmentResults.CurrentPage)
            {
                vm.appointments.Add(appointment);
            }
           
             return PartialView("_CheckAppointments", vm);
        }

        [HttpGet]
        public async Task<ActionResult> LeadAppointment(string leadID)
        {
            var client = await GetExchangeClient();
            LeadAppointmentViewModel vm = new LeadAppointmentViewModel();

            vm.leadID = leadID;
            
            //look up lead information

            //get authorization token
            Authenticator authenticator = new Authenticator();
            var authInfo = await authenticator.AuthenticateAsync(ExchangeResourceId);

            HttpClient httpClient = new HttpClient();
            httpClient.BaseAddress = new Uri(ExchangeResourceId);
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await authInfo.GetAccessToken());

            var response = httpClient.GetAsync("/ews/odata/Me/Inbox/Messages('" + leadID + "')").Result;

            if (response.StatusCode == HttpStatusCode.OK)
            {
                var responseContent = await response.Content.ReadAsStringAsync();
                var message = JObject.Parse(responseContent).ToObject<Message>();
                
                vm.appointmentMessage = message.BodyPreview;
            }

            return View(vm);
        }

        [HttpPost]
        public async Task<ActionResult> LeadAppointment(LeadAppointmentViewModel vm)
        {
            var client = await GetExchangeClient();

            Event appointment = new Event();
            appointment.Subject = "Sales Lead Meeting";
            appointment.Start = new DateTimeOffset(vm.appointmentDate);
            appointment.End = new DateTimeOffset(vm.appointmentDate.AddMinutes(30));
            appointment.BodyPreview = vm.appointmentMessage;            

            await client.Me.Events.AddEventAsync(appointment);

            return RedirectToAction("Index");
        }
	}
	
}