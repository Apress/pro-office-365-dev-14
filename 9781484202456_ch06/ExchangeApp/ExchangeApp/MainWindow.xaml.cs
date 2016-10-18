using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Web;
using Microsoft.Exchange.WebServices.Data;
using System.Windows.Threading;

namespace ExchangeApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ExchangeDataContext _context;
        private PullSubscription _pullSubscription;
        private StreamingSubscription _streamSubscription;
        DispatcherTimer _timer;

        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            _context = new ExchangeDataContext(txtEmailAddress.Text, txtPassword.Password);
            lblOrganizer.Text += " " + txtEmailAddress.Text; 
            EnableButtons();
        }

        protected void EnableButtons()
        {
            btnGetAvailability.IsEnabled = true;
            btnGetItems.IsEnabled = true;
            btnPullSubscribe.IsEnabled = true;
            btnStreamSubscribe.IsEnabled = true;
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BindWellKnownFolderList();
        }

        protected void BindWellKnownFolderList()
        {
            //bind the known folders to the list box
            foreach (WellKnownFolderName folderName in
                Enum.GetValues(typeof(WellKnownFolderName)))
            {
                cboWellKnownFolders.Items.Add(folderName);
            }
        }

        private void btnGetItems_Click(object sender, RoutedEventArgs e)
        {
            //check for given values
            if (string.IsNullOrWhiteSpace(txtEmailAddress.Text))
            {
                MessageBox.Show("You must enter an email address to proceed.");
                return;
            }

            if (cboWellKnownFolders.SelectedIndex < 0)
            {
                MessageBox.Show("You must select a folder to proceed");
                return;
            }

            //get items for the given folder and bind them to the list box
            lstItems.ItemsSource = _context.GetMailboxItems
                ((WellKnownFolderName)cboWellKnownFolders.SelectedItem);
        }

        private void lstItems_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Item email = _context.GetItem(((Item)lstItems.SelectedItem).Id);

            txtMessageBody.Text = "From:" + ((EmailMessage)email).Sender
                                          + Environment.NewLine;
            txtMessageBody.Text += Environment.NewLine;
            txtMessageBody.Text += email.Body;
        }

        private void btnGetAvailability_Click(object sender, RoutedEventArgs e)
        {
            List<string> attendees = new List<string>();
            if (!string.IsNullOrWhiteSpace(txtAttendee1.Text))
            {
                attendees.Add(txtAttendee1.Text);
            }

            if (!string.IsNullOrWhiteSpace(txtAttendee2.Text))
            {
                attendees.Add(txtAttendee2.Text);
            }

            if (attendees.Count == 0)
            {
                MessageBox.Show("You must add at least one attendee to proceed.");
                return;
            }

            GetUserAvailabilityResults results =
                _context.GetAvailability(txtEmailAddress.Text, attendees, 30, 2);

            foreach (Suggestion suggestion in results.Suggestions)
            {
                foreach (TimeSuggestion time in suggestion.TimeSuggestions)
                {
                    lstSuggestions.Items.Add(time);
                }
            }
        }

        private void btnPullSubscribe_Click(object sender, RoutedEventArgs e)
        {
            ExchangeService service = _context.GetService();

            _pullSubscription = service.SubscribeToPullNotifications
                (new FolderId[] { WellKnownFolderName.Inbox }, 10, null, EventType.NewMail);
            txtSubscriptionActivity.Text
                += "Pull Subscription Created" + Environment.NewLine;

            //set up polling
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(10);
            _timer.Tick += timer_Tick;
            _timer.Start();

            btnPullSubscribe.IsEnabled = false;
            btnPullUnsubscribe.IsEnabled = true;
        }

        private void btnPullUnsubscribe_Click(object sender, RoutedEventArgs e)
        {
            _timer.Stop();

            _pullSubscription.Unsubscribe();

            txtSubscriptionActivity.Text
                += "Pull Subscription Unsubscribed" + Environment.NewLine;

            btnPullSubscribe.IsEnabled = true;
            btnPullUnsubscribe.IsEnabled = false;
        }

        private void btnStreamSubscribe_Click(object sender, RoutedEventArgs e)
        {
            ExchangeService service = _context.GetService();

            _streamSubscription = service.SubscribeToStreamingNotifications
                (new FolderId[] { WellKnownFolderName.Inbox }, EventType.NewMail);

            StreamingSubscriptionConnection connection =
                new StreamingSubscriptionConnection(service, 10);

            connection.AddSubscription(_streamSubscription);
            connection.OnNotificationEvent += connection_OnNotificationEvent;

            connection.Open();

            txtSubscriptionActivity.Text
                += "Stream Subscription Created" + Environment.NewLine;

            btnStreamSubscribe.IsEnabled = false;
            btnStreamUnsubscribe.IsEnabled = true;
        }

        private void btnStreamUnsubscribe_Click(object sender, RoutedEventArgs e)
        {
            _streamSubscription.Unsubscribe();

            txtSubscriptionActivity.Text
                += "Stream Subscription Unsubscribed" + Environment.NewLine;

            btnStreamSubscribe.IsEnabled = true;
            btnStreamUnsubscribe.IsEnabled = false;
        }

        void timer_Tick(object sender, EventArgs e)
        {
            GetEventsResults results = _pullSubscription.GetEvents();

            txtSubscriptionActivity.Text
                += "Pull Subscription checked for new items" + Environment.NewLine;

            foreach (ItemEvent itemEvent in results.ItemEvents)
            {
                switch (itemEvent.EventType)
                {
                    case EventType.NewMail:
                        txtSubscriptionActivity.Text
                            += "Pull Subscription: New email received"
                            + Environment.NewLine;
                        break;
                }
            }
        }

        void connection_OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            foreach (NotificationEvent notification in args.Events)
            {
                switch (notification.EventType)
                {
                    case EventType.NewMail:
                        Dispatcher.Invoke(new Action(
                            delegate()
                            {
                                txtSubscriptionActivity.Text
                                    += "Stream Subscription: New email received"
                                    + Environment.NewLine;
                            }));
                        break;
                }
            }
        }

    }
}
