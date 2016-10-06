using RepairShoprCore;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.SQLite;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RepairShoprApps
{
    internal delegate bool CheckFoCancellation();
    internal delegate void ReportCustomerErrorEventHandler(CommitCRM.Account account, string message, Exception exception);
    internal delegate void ReportContactErrorEventHandler(CommitCRM.Contact contact, string message, Exception exception);
    internal delegate void ReportProgressEventHandler(int index, double percentage, bool cancelled);
    internal delegate void ReportStatusEventHandler(string message);

    internal class CustomerExporter
    {
        private const int DATE_RANGE_TO_LOAD = 1;
        private const int ITEMS_COUNT_TO_LOAD = 20;
        private const int MAX_THREADS_COUNT = 10;

        private int _index;
        private int _customerCount;
        private string _connectionString;
        private CheckFoCancellation _checkFoCancellation;
        private ICollection<IEnumerable<CommitCRM.Account>> _customerGroups;

        public event ReportCustomerErrorEventHandler ReportCustomerErrorEvent;
        public event ReportContactErrorEventHandler ReportContactErrorEvent;
        public event ReportProgressEventHandler ReportProgressEvent;
        public event ReportStatusEventHandler ReportStatusEvent;

        public CustomerExporter(string connectionString, CheckFoCancellation checkFoCancellation)
        {
            _connectionString = connectionString;
            _checkFoCancellation = checkFoCancellation;
            _customerGroups = new List<IEnumerable<CommitCRM.Account>>();
        }

        public async Task Export(DateTime fromDate)
        {
            DateTime startDate = fromDate;
            DateTime endDate = startDate.AddMonths(DATE_RANGE_TO_LOAD);

            while (startDate < DateTime.Today)
            {
                _index = 0;
                _customerCount = 0;
                _customerGroups.Clear();

                if (_checkFoCancellation())
                {
                    ReportProgressEvent?.Invoke(_index, 100, true);
                    return;
                }

                var loadCustomers = true;
                while (loadCustomers)
                {
                    loadCustomers = await LoadCustomers(startDate, endDate);
                }

                var tasks = new List<Task>();
                foreach (var group in _customerGroups)
                {
                    var task = ExportMultipleCustomers(group);
                    tasks.Add(task);

                    if (tasks.Count == MAX_THREADS_COUNT)
                    {
                        await Task.WhenAll(tasks);
                        tasks.Clear();
                    }
                }

                await Task.WhenAll(tasks);

                Properties.Settings.Default.CustomerExport = startDate;
                Properties.Settings.Default.Save();

                startDate = endDate;
                endDate = startDate.AddMonths(DATE_RANGE_TO_LOAD);
            }
        }

        private async Task<bool> LoadCustomers(DateTime startDate, DateTime endDate)
        {
            return await Task.Factory.StartNew(() =>
            {
                var result = false;

                if (_checkFoCancellation())
                {
                    ReportProgressEvent?.Invoke(_index, 100, true);
                    return result;
                }

                var message = string.Format("Loading Accounts from {0:y} to {1:y}. Loaded {2} accounts.", startDate, endDate, _customerCount);
                ReportStatusEvent?.Invoke(message);
                ReportProgressEvent?.Invoke(_index, 0, false);

                using (var connection = new SQLiteConnection(_connectionString))
                {
                    connection.Open();

                    var accounts = new CommitCRM.ObjectQuery<CommitCRM.Account>(CommitCRM.LinkEnum.linkAND, ITEMS_COUNT_TO_LOAD);
                    accounts.AddCriteria(CommitCRM.Account.Fields.CreationDate, CommitCRM.OperatorEnum.opGreaterThan, startDate);
                    accounts.AddCriteria(CommitCRM.Account.Fields.CreationDate, CommitCRM.OperatorEnum.opLessThan, endDate);
                    accounts.AddSortExpression(CommitCRM.Account.Fields.CreationDate, CommitCRM.SortDirectionEnum.sortASC);
                    foreach (var item in _customerGroups.SelectMany(x => x))
                        accounts.AddCriteria(CommitCRM.Account.Fields.AccountREC_ID, CommitCRM.OperatorEnum.opNotLike, item.AccountREC_ID.ToString());

                    var accountsResult = accounts.FetchObjects();
                    if (accountsResult != null && accountsResult.Any())
                    {
                        _customerCount += accountsResult.Count;
                        _customerGroups.Add(accountsResult);
                    }

                    result = accountsResult.Count == ITEMS_COUNT_TO_LOAD;
                }

                message = string.Format("Loading Accounts from {0:y} to {1:y}. Loaded {2} accounts.", startDate, endDate, _customerCount);
                ReportStatusEvent?.Invoke(message);
                ReportProgressEvent?.Invoke(_index, 0, false);

                return result;
            });
        }

        private async Task ExportMultipleCustomers(IEnumerable<CommitCRM.Account> accounts)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                foreach (CommitCRM.Account account in accounts)
                {
                    try
                    {
                        if (_checkFoCancellation())
                        {
                            ReportProgressEvent?.Invoke(_index, 100, true);
                            return;
                        }

                        string csutomerId = string.Empty;
                        using (SQLiteCommand cmdItemAlready = new SQLiteCommand(string.Format("SELECT CustomerId FROM Account WHERE AccountId='{0}'", account.AccountREC_ID), connection))
                        {
                            using (SQLiteDataReader reader = cmdItemAlready.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    csutomerId = reader[0].ToString();
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(csutomerId))
                        {
                            _index++;
                            continue;
                        }

                        var message = string.Format("Exporting ( {0}/{1} ) of Account", _index, _customerCount);
                        ReportStatusEvent?.Invoke(message);
                        ReportProgressEvent?.Invoke(_index, (100 * _index) / _customerCount, false);

                        await ExportSingleCustomer(account, connection);

                        ReportProgressEvent?.Invoke(_index, (100 * _index) / _customerCount, false);
                    }
                    catch (Exception ex)
                    {
                        var message = string.Format("Failed to Export New Contact.Due to  { 0}", ex.Message);
                        ReportCustomerErrorEvent?.Invoke(account, message, ex);
                    }
                    _index++;
                }
            }
        }

        private async Task<Customer> ExportSingleCustomer(CommitCRM.Account account, SQLiteConnection connection)
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

                    var contactSearch = new CommitCRM.ObjectQuery<CommitCRM.Contact>();
                    contactSearch.AddCriteria(CommitCRM.Contact.Fields.ParentAccountREC_ID, CommitCRM.OperatorEnum.opEqual, account.AccountREC_ID);
                    contactSearch.AddCriteria(CommitCRM.Contact.Fields.AccountType, CommitCRM.OperatorEnum.opEqual, 5);
                    List<CommitCRM.Contact> contacts = contactSearch.FetchObjects();

                    foreach (CommitCRM.Contact contact in contacts)
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

                        var result = RepairShoprUtils.ExportContact(contactNameCollection);
                        if (result == null)
                        {
                            var message = string.Format("Unable to create Contact with Name : {0}", contactname);
                            ReportContactErrorEvent?.Invoke(contact, message, null);
                        }
                    }
                }
                else
                {
                    var message = string.Format("Unable to create Account with Name : {0}", fullname);
                    ReportCustomerErrorEvent?.Invoke(account, message, null);
                }

                return newCustomer;
            });
        }
    }
}