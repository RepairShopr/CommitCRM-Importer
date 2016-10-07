using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;

namespace RepairShoprApps
{
    internal class TicketExporter : BaseExporter
    {
        private const int DATE_RANGE_TO_LOAD = 1;
        private const int ITEMS_COUNT_TO_LOAD = 10;
        private const int MAX_THREADS_COUNT = 10;

        private readonly int? _defaultLocationId;
        private readonly string _connectionString;

        private int _index;
        private int _ticketCount;
        private CheckFoCancellation _checkFoCancellation;
        private ICollection<IEnumerable<CommitCRM.Ticket>> _ticketGroups;

        public event ReportTicketErrorEventHandler ReportTicketErrorEvent;
        public event ReportProgressEventHandler ReportProgressEvent;
        public event ReportStatusEventHandler ReportStatusEvent;

        public TicketExporter(int? defaultLocationId, string connectionString, CheckFoCancellation checkFoCancellation)
        {
            _defaultLocationId = defaultLocationId;
            _connectionString = connectionString;
            _checkFoCancellation = checkFoCancellation;
            _ticketGroups = new List<IEnumerable<CommitCRM.Ticket>>();
        }

        public async Task Export(DateTime fromDate)
        {
            DateTime startDate = fromDate;
            DateTime endDate = startDate.AddMonths(DATE_RANGE_TO_LOAD);

            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                while (startDate < DateTime.Today)
                {
                    _index = 0;
                    _ticketCount = 0;
                    _ticketGroups.Clear();

                    if (_checkFoCancellation())
                    {
                        ReportProgressEvent?.Invoke(_index, 100, true);
                        connection.Close();
                        return;
                    }

                    var loadTickets = true;
                    while (loadTickets)
                    {
                        loadTickets = await LoadTickets(startDate, endDate);
                    }

                    var tasks = new List<Task>();
                    foreach (var group in _ticketGroups)
                    {
                        var task = ExportMultipleTickets(group, connection);
                        tasks.Add(task);

                        if (tasks.Count == MAX_THREADS_COUNT)
                        {
                            await Task.WhenAll(tasks);
                            tasks.Clear();
                        }
                    }

                    await Task.WhenAll(tasks);

                    Properties.Settings.Default.TicketExport = startDate;
                    Properties.Settings.Default.Save();

                    startDate = endDate;
                    endDate = startDate.AddMonths(DATE_RANGE_TO_LOAD);
                }

                connection.Close();
            }
        }

        private async Task<bool> LoadTickets(DateTime startDate, DateTime endDate)
        {
            return await Task.Factory.StartNew(() =>
            {
                if (_checkFoCancellation())
                {
                    ReportProgressEvent?.Invoke(_index, 100, true);
                    return false;
                }

                var message = string.Format("Loading Tickets from {0:y} to {1:y}. Loaded {2} tickets.", startDate, endDate, _ticketCount);
                ReportStatusEvent?.Invoke(message);
                ReportProgressEvent?.Invoke(_index, 0, false);

                var tickets = new CommitCRM.ObjectQuery<CommitCRM.Ticket>(CommitCRM.LinkEnum.linkAND, ITEMS_COUNT_TO_LOAD);
                tickets.AddCriteria(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.OperatorEnum.opGreaterThan, startDate);
                tickets.AddCriteria(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.OperatorEnum.opLessThan, endDate);
                tickets.AddSortExpression(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.SortDirectionEnum.sortASC);
                foreach (var item in _ticketGroups.SelectMany(x => x))
                    tickets.AddCriteria(CommitCRM.Ticket.Fields.TicketREC_ID, CommitCRM.OperatorEnum.opNotLike, item.TicketREC_ID.ToString());

                var ticketResult = tickets.FetchObjects();
                if (ticketResult != null && ticketResult.Any())
                {
                    _ticketCount += ticketResult.Count;
                    _ticketGroups.Add(ticketResult);
                }

                return ticketResult.Count == ITEMS_COUNT_TO_LOAD;
            });
        }

        private async Task ExportMultipleTickets(IEnumerable<CommitCRM.Ticket> tickets, SQLiteConnection connection)
        {
            foreach (CommitCRM.Ticket ticket in tickets)
            {
                try
                {
                    if (_checkFoCancellation())
                    {
                        ReportProgressEvent?.Invoke(_index, 100, true);
                        return;
                    }

                    string ticketId = string.Empty;
                    using (SQLiteCommand cmdItemAlready = new SQLiteCommand(string.Format("SELECT RTicketId FROM Ticket WHERE ticketId='{0}'", ticket.TicketREC_ID), connection))
                    {
                        using (SQLiteDataReader reader = cmdItemAlready.ExecuteReader())
                        {
                            while (reader.Read())
                                ticketId = reader[0].ToString();
                        }
                    }

                    if (!string.IsNullOrEmpty(ticketId))
                    {
                        _index++;
                        continue;
                    }

                    string customerId = string.Empty;
                    using (SQLiteCommand cmdItemAlready = new SQLiteCommand(string.Format("SELECT CustomerId FROM Account WHERE AccountId='{0}'", ticket.AccountREC_ID), connection))
                    {
                        using (SQLiteDataReader reader = cmdItemAlready.ExecuteReader())
                        {
                            while (reader.Read())
                                customerId = reader[0].ToString();
                        }
                    }

                    if (string.IsNullOrEmpty(customerId))
                    {
                        var customer = await CreateCustomerForTicket(ticket, connection);
                        if (customer != null)
                            customerId = customer.Id;
                        else
                        {
                            var customerIdMessage = string.Format("Unable to create Customer of Ticket: {0}", ticket.Description);
                            ReportTicketErrorEvent?.Invoke(ticket, customerIdMessage, null);
                            _index++;
                            continue;
                        }
                    }

                    var message = string.Format("Exporting ( {0}/{1} ) of Ticket", _index, _ticketCount);
                    ReportStatusEvent?.Invoke(message);
                    ReportProgressEvent?.Invoke(_index, (100 * _index) / _ticketCount, false);

                    var result = await ExportSingleTicket(ticket, _defaultLocationId, customerId, connection);
                    if (result == null)
                    {
                        message = string.Format("Unable to create Ticket: {0}", ticket.Description);
                        ReportTicketErrorEvent?.Invoke(ticket, message, null);
                    }
                }
                catch (Exception ex)
                {
                    var message = string.Format("Failed to Export a New Ticket {1}. Due to {0}", ex.Message, ticket.Description);
                    ReportTicketErrorEvent?.Invoke(ticket, message, ex);
                }
                _index++;
            }
        }
    }
}