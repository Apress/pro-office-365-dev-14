using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.Windows;

namespace ExchangeApp
{
    class ExchangeDataContext
    {
        private ExchangeService _service;

        public ExchangeDataContext(string emailAddress, string password)
        {
            _service = GetBinding(emailAddress, password);
        }

        public ExchangeService GetService()
        {
            return _service;
        }

        static ExchangeService GetBinding(string emailAddress, string password)
        {
            // Create the binding.
            ExchangeService service =
                new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            // Define credentials.
            service.Credentials = new WebCredentials(emailAddress, password);

            // Use the AutodiscoverUrl method to locate the service endpoint.
            try
            {
                service.AutodiscoverUrl(emailAddress, RedirectionUrlValidationCallback);
            }
            catch (AutodiscoverRemoteException ex)
            {
                MessageBox.Show("Autodiscover error: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

            return service;
        }

        static bool RedirectionUrlValidationCallback(String redirectionUrl)
        {
            // Perform validation.
            // Validation is developer dependent to ensure a safe redirect.
            return true;
        }

        public List<Folder> GetFolders(FolderId parentFolderID)
        {
            return _service.FindFolders(parentFolderID, null).ToList();
        }

        public List<Item> GetMailboxItems(WellKnownFolderName folder)
        {
            return _service.FindItems(folder, new ItemView(30)).ToList();
        }

        public Item GetItem(ItemId itemId)
        {
            List<ItemId> items = new List<ItemId>() { itemId };

            PropertySet properties = new PropertySet(BasePropertySet.IdOnly,
                EmailMessageSchema.Body, EmailMessageSchema.Sender,
                EmailMessageSchema.Subject);
            properties.RequestedBodyType = BodyType.Text;
            ServiceResponseCollection<GetItemResponse> response =
                _service.BindToItems(items, properties);

            return response[0].Item;
        }

        public GetUserAvailabilityResults GetAvailability
            (string organizer,
             List<string> requiredAttendees,
             int meetingDuration,
             int timeWindowDays)
        {
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();

            //add organizer
            attendees.Add(new AttendeeInfo()
            {
                SmtpAddress = organizer,
                AttendeeType = MeetingAttendeeType.Organizer
            });

            //add required attendees
            foreach (string attendee in requiredAttendees)
            {
                attendees.Add(new AttendeeInfo()
                {
                    SmtpAddress = attendee,
                    AttendeeType = MeetingAttendeeType.Required
                });
            }

            //setup options
            AvailabilityOptions options = new AvailabilityOptions()
            {
                MeetingDuration = meetingDuration,
                MaximumNonWorkHoursSuggestionsPerDay = 4,
                MinimumSuggestionQuality = SuggestionQuality.Good,
                RequestedFreeBusyView = FreeBusyViewType.FreeBusy
            };

            GetUserAvailabilityResults results = _service.GetUserAvailability
                (attendees,
                 new TimeWindow(DateTime.Now, DateTime.Now.AddDays(timeWindowDays)),
                 AvailabilityData.FreeBusyAndSuggestions,
                 options);

            return results;
        }

    }
}
