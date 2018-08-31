using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace RepairShoprApps
{
    internal class CustomerExporter : BaseExporter
    {
        private const int DATE_RANGE_TO_LOAD = 1;
        private const int ITEMS_COUNT_TO_LOAD = 10;
        private const int MAX_THREADS_COUNT = 10;

        private readonly string _connectionString;

        private int _index;
        private int _customerCount;
        private CheckFoCancellation _checkFoCancellation;
        private ICollection<IEnumerable<CommitCRM.Account>> _customerGroups;

        public event ReportCustomerErrorEventHandler ReportCustomerErrorEvent;
        public event ReportProgressEventHandler ReportProgressEvent;
        public event ReportStatusEventHandler ReportStatusEvent;

        public CustomerExporter(string connectionString, CheckFoCancellation checkFoCancellation)
        {
            _connectionString = connectionString;
            _checkFoCancellation = checkFoCancellation;
            _customerGroups = new List<IEnumerable<CommitCRM.Account>>();
        }

        public async Task Export(DateTime fromDate, string remoteHost)
        {
            DateTime startDate = fromDate;
            DateTime endDate = startDate.AddMonths(DATE_RANGE_TO_LOAD);

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                while (startDate < DateTime.Today)
                {
                    _index = 0;
                    _customerCount = 0;
                    _customerGroups.Clear();

                    if (_checkFoCancellation())
                    {
                        ReportProgressEvent?.Invoke(_index, 100, true);
                        connection.Close();
                        return;
                    }

                    var loadCustomers = true;
                    while (loadCustomers)
                    {
                        //divide in groups of 10
                        loadCustomers = await LoadCustomers(startDate, endDate);
                    }

                    var tasks = new List<Task>();
                    foreach (var group in _customerGroups)
                    {
                        var task = ExportMultipleCustomers(group, connection, remoteHost);
                        tasks.Add(task);

                        if (tasks.Count == MAX_THREADS_COUNT)
                        {
                            await Task.WhenAll(tasks);
                            tasks.Clear();
                        }
                    }

                    await Task.WhenAll(tasks);

                    startDate = endDate;
                    endDate = startDate.AddMonths(DATE_RANGE_TO_LOAD);
                }

                connection.Close();
            }
        }

        private async Task<bool> LoadCustomers(DateTime startDate, DateTime endDate)
        {
            return await Task.Factory.StartNew(() =>
            {
                if (_checkFoCancellation())
                {
                    ReportProgressEvent?.Invoke(_index, 100, true);
                    return false;
                }

                var message = string.Format("Loading Accounts from {0:y} to {1:y}. Loaded {2} accounts.", startDate, endDate, _customerCount);
                ReportStatusEvent?.Invoke(message);
                ReportProgressEvent?.Invoke(_index, 0, false);

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

                return accountsResult.Count == ITEMS_COUNT_TO_LOAD;
            });
        }

        private async Task ExportMultipleCustomers(IEnumerable<CommitCRM.Account> accounts, SQLiteConnection connection, string remoteHost)
        {   
            Debug.WriteLine(accounts);
            foreach (CommitCRM.Account account in accounts)
            {
                try
                {
                    if (_checkFoCancellation())
                    {
                        ReportProgressEvent?.Invoke(_index, 100, true);
                        return;
                    }

                    string customerId = string.Empty;
                    //check cache and skip already exported
                    using (SQLiteCommand cmdItemAlready = new SQLiteCommand(string.Format("SELECT CustomerId FROM Account WHERE AccountId='{0}'", account.AccountREC_ID), connection))
                    {
                        using (SQLiteDataReader reader = cmdItemAlready.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                Debug.WriteLine(reader[0]);
                                customerId = reader[0].ToString();
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(customerId))
                    {
                        _index++;
                        continue;
                    }

                    var message = string.Format("Exporting ( {0}/{1} ) of Account", _index, _customerCount);
                    ReportStatusEvent?.Invoke(message);
                    ReportProgressEvent?.Invoke(_index, (100 * _index) / _customerCount, false);

                    var customer = await ExportSingleCustomer(account, connection, remoteHost);
                    if (customer == null)
                    {
                        message = string.Format("Unable to create Account with Name : {0}", account.LastName);
                        ReportCustomerErrorEvent?.Invoke(account, message, null);
                    }
                }
                catch (Exception ex)
                {
                    var message = string.Format("Failed to Export a New Account {1}. Due to {0}", ex.Message, account.LastName);
                    ReportCustomerErrorEvent?.Invoke(account, message, ex);
                }
                _index++;
            }
        }
    }
}