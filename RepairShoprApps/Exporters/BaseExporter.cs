using RepairShoprCore;
using System;
using System.Collections.Specialized;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace RepairShoprApps
{
    internal delegate bool CheckFoCancellation();
    internal delegate void ReportProgressEventHandler(int index, double percentage, bool cancelled);
    internal delegate void ReportStatusEventHandler(string message);
    internal delegate void ReportCustomerErrorEventHandler(CommitCRM.Account account, string message, Exception exception);
    internal delegate void ReportTicketErrorEventHandler(CommitCRM.Ticket ticket, string message, Exception exception);

    internal abstract class BaseExporter
    {
        #region Methods
        protected async Task<Customer> ExportSingleCustomer(CommitCRM.Account account, SQLiteConnection connection)
        {
            return await Task.Factory.StartNew(() =>
            {
                var myNameValueCollection = new NameValueCollection();
                string fullname = account.GetFieldValue("FLDCRDCONTACT");
                myNameValueCollection.Add("business_name", account.CompanyName);
                if (!string.IsNullOrEmpty(fullname) && !string.IsNullOrEmpty(account.LastName))
                    myNameValueCollection.Add("firstname", fullname.Replace(account.LastName, string.Empty));
                else
                    myNameValueCollection.Add("firstname", fullname);
                myNameValueCollection.Add("lastname", account.LastName);
                myNameValueCollection.Add("email", account.EmailAddress1);
                myNameValueCollection.Add("phone", Regex.Replace(account.Phone1, @"[^.0-9\s]", ""));
                myNameValueCollection.Add("mobile", Regex.Replace(account.Phone2, @"[^.0-9\s]", ""));
                myNameValueCollection.Add("address", account.AddressLine1);
                myNameValueCollection.Add("address_2", account.AddressLine2);
                myNameValueCollection.Add("city", account.City);
                myNameValueCollection.Add("state", account.State);
                myNameValueCollection.Add("zip", account.Zip);
                myNameValueCollection.Add("notes", account.Notes);

                var newCustomer = RepairShoprUtils.ExportCustomer(myNameValueCollection);
                if (newCustomer != null)
                {
                    using (SQLiteCommand cmdINewItem = new SQLiteCommand(string.Format("INSERT INTO  Account (AccountId,CustomerId) VALUES('{0}','{1}')", account.AccountREC_ID, newCustomer.Id), connection))
                        cmdINewItem.ExecuteNonQuery();

                    var contacts = new CommitCRM.ObjectQuery<CommitCRM.Contact>();
                    contacts.AddCriteria(CommitCRM.Contact.Fields.ParentAccountREC_ID, CommitCRM.OperatorEnum.opEqual, account.AccountREC_ID);
                    contacts.AddCriteria(CommitCRM.Contact.Fields.AccountType, CommitCRM.OperatorEnum.opEqual, 5);

                    var contactsResult = contacts.FetchObjects();
                    foreach (CommitCRM.Contact contact in contactsResult)
                    {
                        string contactname = contact.GetFieldValue("FLDCRDCONTACT");
                        NameValueCollection contactNameCollection = new NameValueCollection();
                        if (account.EmailAddress1.Contains("@"))
                            myNameValueCollection.Add("email", contact.EmailAddress1);
                        contactNameCollection.Add("phone", contact.Phone1);
                        contactNameCollection.Add("mobile", contact.Phone2);
                        contactNameCollection.Add("address", contact.AddressLine1);
                        contactNameCollection.Add("address_2", contact.AddressLine2);
                        contactNameCollection.Add("city", contact.City);
                        contactNameCollection.Add("state", contact.State);
                        contactNameCollection.Add("zip", contact.Zip);
                        contactNameCollection.Add("customer_id", newCustomer.Id);
                        contactNameCollection.Add("name", contactname);

                        RepairShoprUtils.ExportContact(contactNameCollection);
                    }
                }

                return newCustomer;
            });
        }

        protected async Task<Ticket> ExportSingleTicket(CommitCRM.Ticket ticket, int? defaultLocationId, string customerId, SQLiteConnection connection)
        {
            return await Task.Factory.StartNew(() =>
            {
                var myNameValueCollection = new NameValueCollection();
                string subject = string.Empty;
                if (!string.IsNullOrEmpty(ticket.Description))
                {
                    if (ticket.Description.Length > 255)
                        subject = ticket.Description.Substring(0, 255);
                    else
                        subject = ticket.Description;
                }

                if (defaultLocationId.HasValue)
                    myNameValueCollection.Add("location_id", defaultLocationId.Value.ToString());
                myNameValueCollection.Add("subject", subject);
                myNameValueCollection.Add("customer_id", customerId);
                if (!string.IsNullOrEmpty(ticket.TicketType))
                    myNameValueCollection.Add("problem_type", ticket.TicketType);
                myNameValueCollection.Add("status", "Resolved");
                myNameValueCollection.Add("comment_body", GetComment(ticket));
                if (!string.IsNullOrEmpty(ticket.Status_Text))
                    myNameValueCollection.Add("comment_subject", ticket.Status_Text);
                myNameValueCollection.Add("comment_hidden", "1");
                myNameValueCollection.Add("comment_do_not_email", "1");
                string createDate = HttpUtility.UrlEncode(ticket.UpdateDate.ToString("yyyy-MM-dd H:mm:ss"));
                myNameValueCollection.Add("created_at", createDate);

                var newTicket = RepairShoprUtils.ExportTicket(myNameValueCollection);
                if (newTicket != null)
                {
                    using (SQLiteCommand cmdINewItem = new SQLiteCommand(string.Format("INSERT INTO  Ticket (TicketId,RTicketId) VALUES('{0}','{1}')", ticket.TicketREC_ID, newTicket.Id), connection))
                        cmdINewItem.ExecuteNonQuery();
                }

                return newTicket;
            });
        }

        protected async Task<Customer> CreateCustomerForTicket(CommitCRM.Ticket ticket, SQLiteConnection connection)
        {
            var accounts = new CommitCRM.ObjectQuery<CommitCRM.Account>(CommitCRM.LinkEnum.linkAND, 1);
            accounts.AddCriteria(CommitCRM.Account.Fields.AccountREC_ID, CommitCRM.OperatorEnum.opEqual, ticket.AccountREC_ID);

            var accountsResult = accounts.FetchObjects();
            var account = accountsResult.SingleOrDefault();

            if (account != null)
                return await ExportSingleCustomer(account, connection);
            else
                return null;
        }

        private string GetComment(CommitCRM.Ticket ticket)
        {
            var sb = new StringBuilder();
            sb.Append(string.Format("Ticket Number: {0}", ticket.TicketNumber.Replace("-", string.Empty)));
            sb.AppendFormat(" , Cause: {0}", ticket.GetFieldValue("FLDTKTCAUSE"));
            sb.AppendFormat(" , Source: {0}", ticket.Source);
            sb.AppendFormat(" , Category: {0}", ticket.GetFieldValue("FLDTKTCATEGORY"));
            sb.AppendFormat(" , Note: {0}", ticket.Notes);
            sb.AppendFormat(" , Resolution : {0}", ticket.Resolution);

            var charges = new CommitCRM.ObjectQuery<CommitCRM.Charge>();
            charges.AddCriteria(CommitCRM.Charge.Fields.TicketREC_ID, CommitCRM.OperatorEnum.opEqual, ticket.TicketREC_ID);

            var chargesResult = charges.FetchObjects();
            foreach (CommitCRM.Charge charge in chargesResult)
            {
                sb.AppendFormat(", Charge :{0}", charge.Description);
                sb.AppendFormat(", Amount : {0}", charge.GetFieldValue("FLDSLPBILLTOTAL"));
                sb.AppendFormat(", Quantity: {0}", charge.GetFieldValue("FLDSLPQUANTITY"));
                sb.AppendFormat(", Date : {0}", charge.Date);
            }

            var history = new CommitCRM.ObjectQuery<CommitCRM.HistoryNote>();
            history.AddCriteria(CommitCRM.HistoryNote.Fields.RelLinkREC_ID, CommitCRM.OperatorEnum.opEqual, ticket.TicketREC_ID);

            var historyResult = history.FetchObjects();
            foreach (CommitCRM.HistoryNote historyNote in historyResult)
            {
                sb.AppendFormat(" History Note :{0}", historyNote.Description);
                sb.AppendFormat(", Date : {0}", historyNote.Date);
            }

            return sb.ToString();
        }
        #endregion
    }
}