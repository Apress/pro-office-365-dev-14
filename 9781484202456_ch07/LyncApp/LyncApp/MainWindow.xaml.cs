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

using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Extensibility;

namespace LyncApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private LyncClient _client = null;
        private Automation _automation = null;
        private string _remoteUri = "";
        private ConversationWindow _conversationWindow;

        private delegate void FocusWindow();
        private delegate void ResizeWindow(Size newSize);

        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;

            _client = LyncClient.GetClient();
            _automation = LyncClient.GetAutomation();
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BuildContactList();
            LoadContacts();
            LoadCustomContacts();
        }

        private List<Contact> _contacts = new List<Contact>();

        protected void BuildContactList()
        {
            // Build collection of valid contacts -- using Contact class
            _contacts.Add(new Contact() 
            { name = "Mark Collins", 
              sipAddress = "sip:mark@apress365.onmicrosoft.com" });
            _contacts.Add(new Contact() 
            { name = "Michael Mayberry", 
              sipAddress = "sip:michael@apress365.onmicrosoft.com" });
            _contacts.Add(new Contact() 
            { name = "Sahil Malik", 
              sipAddress = "sip:sahilmalik@winsmarts.com" });
        }

        protected void LoadContacts()
        {
            // Bind collection to combo box
            cboContacts.ItemsSource = _contacts;
            cboContacts.DisplayMemberPath = "name";
            cboContacts.SelectedValuePath = "sipAddress";
        }

        protected void LoadCustomContacts()
        {
            lyncCustomList.ItemsSource = (from C in _contacts select C.sipAddress);
        }

        private void cboContacts_SelectionChanged
            (object sender, SelectionChangedEventArgs e)
        {
            // Set the sipAddress of the selected item as the 
            // source for the presence indicator
            lyncPresence.Source = cboContacts.SelectedValue;

            // Set the start instant message button
            lyncStartMessage.Source = cboContacts.SelectedValue;
        }

        private void btnConversationStart_Click(object sender, RoutedEventArgs e)
        {
            _remoteUri = cboContacts.SelectedValue.ToString();

            ConversationManager conversationManager = _client.ConversationManager;
            conversationManager.ConversationAdded 
                += new EventHandler<ConversationManagerEventArgs>
                    (conversationManager_ConversationAdded);
            Conversation conversation = conversationManager.AddConversation();
        }

        private void conversationManager_ConversationAdded(object sender, 
            ConversationManagerEventArgs e)
        {
            e.Conversation.ParticipantAdded 
                += new EventHandler<ParticipantCollectionChangedEventArgs>
                    (Conversation_ParticipantAdded);
            e.Conversation.AddParticipant
                (_client.ContactManager.GetContactByUri(_remoteUri));

            _conversationWindow = _automation.GetConversationWindow(e.Conversation);

            //wire up events
            _conversationWindow.NeedsSizeChange += _conversationWindow_NeedsSizeChange;
            _conversationWindow.NeedsAttention += _conversationWindow_NeedsAttention;

            //dock conversation window
            _conversationWindow.Dock(formHost.Handle);
        }

        private void Conversation_ParticipantAdded(object sender, 
            ParticipantCollectionChangedEventArgs e)
        {
            // add event handlers for modalities of participants:
            if (e.Participant.IsSelf == false)
            {
                if (((Conversation)sender)
                    .Modalities.ContainsKey(ModalityTypes.InstantMessage))
                {
                    ((InstantMessageModality)e.Participant
                        .Modalities[ModalityTypes.InstantMessage]).InstantMessageReceived 
                        += new EventHandler<MessageSentEventArgs>
                            (ConversationTest_InstantMessageReceived);

                    ((InstantMessageModality)e.Participant
                        .Modalities[ModalityTypes.InstantMessage]).IsTypingChanged 
                        += new EventHandler<IsTypingChangedEventArgs>
                            (ConversationTest_IsTypingChanged);
                }
                Conversation conversation = (Conversation)sender;

                InstantMessageModality imModality = 
                    (InstantMessageModality)conversation
                    .Modalities[ModalityTypes.InstantMessage];

                IDictionary<InstantMessageContentType, string> textMessage = 
                    new Dictionary<InstantMessageContentType, string>();
                textMessage.Add(InstantMessageContentType.PlainText, "Hello, World!");

                if (imModality.CanInvoke(ModalityAction.SendInstantMessage))
                {
                    IAsyncResult asyncResult = imModality.BeginSendMessage(
                        textMessage,
                        SendMessageCallback,
                        imModality);
                }
            }
        }

        private void SendMessageCallback(IAsyncResult ar)
        {
            InstantMessageModality imModality = (InstantMessageModality)ar.AsyncState;

            try
            {
                imModality.EndSendMessage(ar);
            }
            catch (LyncClientException lce)
            {
                MessageBox.Show("Lync Client Exception on EndSendMessage " + lce.Message);
            }

        }

        private void ConversationTest_IsTypingChanged(object sender, 
            IsTypingChangedEventArgs e)
        {

        }

        private void ConversationTest_InstantMessageReceived(object sender, 
            MessageSentEventArgs e)
        {

        }

        private void _conversationWindow_NeedsAttention(object sender, 
            ConversationWindowNeedsAttentionEventArgs e)
        {
            FocusWindow focusWindow = new FocusWindow(GetWindowFocus);
            Dispatcher.Invoke(focusWindow, new object[] { });
        }

        private void _conversationWindow_NeedsSizeChange(object sender, 
            ConversationWindowNeedsSizeChangeEventArgs e)
        {
            Size windowSize = new Size();
            windowSize.Height = e.RecommendedWindowHeight;
            windowSize.Width = e.RecommendedWindowWidth;
            ResizeWindow resize = new ResizeWindow(SetWindowSize);
            Dispatcher.Invoke(resize, new object[] { windowSize });
        }

        private void SetWindowSize(Size newSize)
        {
            formPanel.Size = new System.Drawing.Size(
            (int)newSize.Width, (int)newSize.Height);
        }

        private void GetWindowFocus()
        {
            Focus();
        }
    }
}
