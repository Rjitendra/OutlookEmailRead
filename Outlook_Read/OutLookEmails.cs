using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

namespace Outlook_Read
{
    public class OutLookEmails
    {
        // hi jitendra
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }
        public string EmailCc { get; set; }
        public string EmailBCc { get; set; }
        public string EmailTo { get; set; }
        public string EmailAttachment { get; set; }
        //public string EmailBCc { get; set; }
        //public string EmailTo { get; set; }
        //public string EmailAttachment { get; set; }
        public static List<OutLookEmails> ReadMailItems()
        {
            List<OutLookEmails> listEmailDetails = new List<OutLookEmails>();
            OutLookEmails emailsDetails;
            ExchangeService exchange = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            try
            {
                exchange.Credentials = new WebCredentials("email", "password");

                // exchange.Url = new Uri("https://outlook.office365.com/owa/bridgetree.com/ews/exchange.asmx");
                exchange.AutodiscoverUrl("email", RedirectionUrlValidationCallback);

                TimeSpan ts = new TimeSpan(0, -1, 0, 0);
                DateTime date = DateTime.Today.AddDays(-1);
                //DateTime.Now.Add(ts);

                // new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date)

                SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date), new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

                if (exchange != null)
                {

                    PropertySet FindItemPropertySet = new PropertySet(BasePropertySet.IdOnly);
                    ItemView view = new ItemView(999);
                    view.PropertySet = FindItemPropertySet;

                    PropertySet GetItemsPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
                    GetItemsPropertySet.RequestedBodyType = BodyType.Text;
                    FindItemsResults<Item> emailMessages = null;

                    emailMessages = exchange.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

                    if (emailMessages.Items.Count > 0)
                    {
                        exchange.LoadPropertiesForItems(emailMessages.Items, GetItemsPropertySet);
                        foreach (EmailMessage message in emailMessages.Items)
                        {
                            if (message.IsRead == false && emailMessages.Items.Count > 0)
                            {
                                emailsDetails = new OutLookEmails();
                                emailsDetails.EmailFrom = message.From.ToString();
                                emailsDetails.EmailSubject = message.Subject;
                                emailsDetails.EmailBody = message.Body.Text;
                                listEmailDetails.Add(emailsDetails);
                                message.IsRead = true;
                                message.Update(ConflictResolutionMode.AlwaysOverwrite);

                            }
                            else
                            {

                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {


                throw (ex);
            }
            return listEmailDetails;

        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
