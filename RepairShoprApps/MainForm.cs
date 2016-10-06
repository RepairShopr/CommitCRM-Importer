using Microsoft.Win32;
using RepairShoprCore;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Web;
using System.Windows.Forms;

namespace RepairShoprApps
{
    public partial class MainForm : Form
    {
        private bool _exportCustomer = false;
        private bool _exportTicket = false;
        private string _statusMessage = string.Empty;
        public string _path = null;
        BackgroundWorker bgw = null;
        string installedLocation = string.Empty;
        bool isCompleteCustomer = false;
        bool isCompleteTicket = false;
        private int? _defaultLocationId;

        public MainForm()
        {
            InitializeComponent();
            label1.Text = string.Empty;
            this.Text = "RepairShoprApps for CommitCRM";
            buttonExport.Enabled = false;
            progressBar1.Visible = false;
            label3.Text = "";

            RepairShoprUtils.LogWriteLineinHTML("Initializing System with Default value", MessageSource.Initialization, "", messageType.Information);
            CommitCRMHandle();
            DBhandler();
        }

        private void CommitCRMHandle()
        {
            try
            {
                string p_name = "CommitCRM";
                string displayName;
                RegistryKey key;
                key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                RepairShoprUtils.LogWriteLineinHTML("Reading CommitCRM Install Location from RegistryKey", MessageSource.Initialization, "", messageType.Information);
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey subkey = key.OpenSubKey(keyName);
                    displayName = subkey.GetValue("DisplayName") as string;
                    if (p_name.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                    {
                        installedLocation = subkey.GetValue("InstallLocation") as string;//return true;
                        RepairShoprUtils.LogWriteLineinHTML(string.Format("CommitCRM Install Directiory is '{0}'", installedLocation), MessageSource.Initialization, "", messageType.Information);

                        RepairShoprUtils.LogWriteLineinHTML("Setting CommitCRM Parameter such as DLL folder and DB Folder", MessageSource.Initialization, "", messageType.Information);
                        CommitCRM.Config config = new CommitCRM.Config();
                        config.AppName = "RepairShopr";
                        config.CommitDllFolder = Path.Combine(installedLocation, "ThirdParty", "UserDev");
                        config.CommitDbFolder = Path.Combine(installedLocation, "db");
                        CommitCRM.Application.Initialize(config);
                        RepairShoprUtils.LogWriteLineinHTML("Successfully configure CommitCRM ", MessageSource.Initialization, "", messageType.Information);
                        break;
                    }
                }

            }
            catch (CommitCRM.Exception exc)
            {
                RepairShoprUtils.LogWriteLineinHTML("Failed to Configure CommitCRM", MessageSource.Initialization, exc.Message, messageType.Error);
            }
        }

        private void DBhandler()
        {
            _path = Path.Combine(RepairShoprUtils.folderPath, "RepairShopr.db3");
            FileInfo fileInfo = new FileInfo(_path);
            if (!fileInfo.Exists)
            {
                SQLiteConnection.CreateFile(_path);
                try
                {
                    using (SQLiteConnection con = new SQLiteConnection("data source=" + _path + ";PRAGMA journal_mode=WAL;"))
                    {
                        con.SetPassword("shyam");
                        con.Open();
                        SQLiteCommand cmdTableAccount = new SQLiteCommand("CREATE TABLE Account (Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT ,AccountId nvarchar(50),CustomerId nvarchar(50))", con);
                        cmdTableAccount.ExecuteNonQuery();
                        SQLiteCommand cmdTableTicket = new SQLiteCommand("CREATE TABLE Ticket (Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,TicketId nvarchar(50),RTicketId nvarchar(30))", con);
                        cmdTableTicket.ExecuteNonQuery();
                        SQLiteCommand cmdTableInvoice = new SQLiteCommand("CREATE TABLE Invoice (Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,InvoiceId nvarchar(50),RinvoiceId nvarchar(50))", con);
                        cmdTableInvoice.ExecuteNonQuery();
                        con.Close();
                    }
                }
                catch (Exception ex)
                {
                    RepairShoprUtils.LogWriteLineinHTML("Failed to Create Datebase Setting", MessageSource.Initialization, ex.StackTrace, messageType.Error);
                }
            }
        }

        private void buttonLogin_Click(object sender, EventArgs e)
        {
            label1.ForeColor = Color.Green;
            label1.Text = "Login in progress";
            if (string.IsNullOrEmpty(textBoxUserName.Text))
            {
                errorProvider1.SetError(textBoxUserName, "User Name is Required field");
                label1.ForeColor = Color.Red;
                label1.Text = "User Name is Required field";
                RepairShoprUtils.LogWriteLineinHTML("User Name is is Required ", MessageSource.Login, "", messageType.Error);
                return;
            }
            else
            {
                errorProvider1.Clear();
                label1.ForeColor = Color.Green;
            }
            if (string.IsNullOrEmpty(textBoxPassWord.Text))
            {
                label1.ForeColor = Color.Red;
                errorProvider2.SetError(textBoxPassWord, "Password is Required field");
                label1.Text = "Password is Required field";
                RepairShoprUtils.LogWriteLineinHTML("Password is Required field ", MessageSource.Login, "", messageType.Error);
                return;
            }
            else
            {
                errorProvider2.Clear();
                label1.ForeColor = Color.Green;
            }
            RepairShoprUtils.LogWriteLineinHTML("Sending User Name and Password for Authentication", MessageSource.Login, "", messageType.Information);
            LoginResponse result = RepairShoprUtils.GetLoginResquest(textBoxUserName.Text.Trim(), textBoxPassWord.Text.Trim()); //"shyam@gmail.com", "shyambct525"
            if (result != null)
            {
                label1.Text = "Login Successful";
                buttonExport.Enabled = true;

                _defaultLocationId = result.DefaulLocationId;
            }
            else
            {
                label1.ForeColor = Color.Red;
                label1.Text = "Failed to Authenticate";

                _defaultLocationId = null;
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            CommitCRM.Application.Terminate();
            this.Close();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            if (bgw == null)
            {
                bgw = new BackgroundWorker();
                bgw.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
                bgw.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
                bgw.WorkerReportsProgress = true;
                bgw.WorkerSupportsCancellation = true;
            }
            _exportTicket = checkBoxExportTicket.Checked;
            _exportCustomer = checkBoxExportCustomer.Checked;

            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressBar1.Enabled = true;

            bgw.RunWorkerAsync();
            buttonStop.Enabled = true;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var connectionString = "data source=" + _path + ";PRAGMA journal_mode=WAL;Password=shyam;";
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                var defaultAccounts = new CommitCRM.ObjectQuery<CommitCRM.Account>(CommitCRM.LinkEnum.linkAND, 1);
                defaultAccounts.AddSortExpression(CommitCRM.Account.Fields.CreationDate, CommitCRM.SortDirectionEnum.sortASC);
                var defaultAccountResult = defaultAccounts.FetchObjects();

                DateTime customerExport = Directory.GetCreationTime(installedLocation);
                DateTime ticketExport = new DateTime();
                if (defaultAccountResult != null && defaultAccountResult.Count > 0)
                {
                    ticketExport = customerExport = defaultAccountResult[0].CreationDate;
                }

                #region Customer Export
                if (_exportCustomer)
                {
                    //if (DefaultAccountResult != null && DefaultAccountResult.Count > 0)
                    //    customerExport = DefaultAccountResult[0].CreationDate;
                    if (Properties.Settings.Default.CustomerExport != null && Properties.Settings.Default.CustomerExport > customerExport)
                        customerExport = Properties.Settings.Default.CustomerExport;

                    var exporter = new CustomerExporter(connectionString, () => ((BackgroundWorker)sender).CancellationPending);
                    exporter.ReportStatusEvent += (message) =>
                    {
                        _statusMessage = message;
                    };
                    exporter.ReportProgressEvent += (index, percentage, cancelled) =>
                    {
                        ((BackgroundWorker)sender).ReportProgress((int)percentage, index);
                        if (cancelled)
                            RepairShoprUtils.LogWriteLineinHTML("Contact Exporting Process is Stoped or Cancelled by User", MessageSource.Customer, "", messageType.Warning);
                    };
                    exporter.ReportContactErrorEvent += (contact, message, exception) =>
                    {
                        RepairShoprUtils.LogWriteLineinHTML(message, MessageSource.Customer, exception?.ToString(), messageType.Error);
                    };
                    exporter.ReportCustomerErrorEvent += (customer, message, exception) =>
                    {
                        RepairShoprUtils.LogWriteLineinHTML(message, MessageSource.Customer, exception?.ToString(), messageType.Error);
                    };

                    var task = exporter.Export(customerExport);
                    task.Wait();

                    isCompleteCustomer = true;
                }
                #endregion

                #region Ticket Export
                int index_ = 1;
                int percentage_ = 1;
                int ticketCount = 0;
                int ticketIndex = 1;

                if (_exportTicket)
                {
                    //string startTicket = string.Empty;
                    //int lastTicket = 0;
                    //int ticketNumber = 0;
                    //CommitCRM.ObjectQuery<CommitCRM.Ticket> DefaultTickets = new CommitCRM.ObjectQuery<CommitCRM.Ticket>(CommitCRM.LinkEnum.linkAND, 1);
                    //DefaultTickets.AddSortExpression(CommitCRM.Ticket.Fields.TicketNumber, CommitCRM.SortDirectionEnum.sortASC);
                    //List<CommitCRM.Ticket> DefaultTicketResult = DefaultTickets.FetchObjects();

                    //if (DefaultTicketResult != null && DefaultTicketResult.Count > 0)
                    //    startTicket = DefaultTicketResult[0].TicketNumber;

                    //var partTicketByPart = startTicket.Split('-');

                    //if (!string.IsNullOrEmpty(Properties.Settings.Default.TicketNumber))
                    //    ticketNumber = int.Parse(Properties.Settings.Default.TicketNumber);
                    //else
                    //{
                    //    if (partTicketByPart.Length == 2)
                    //    {
                    //        ticketNumber = int.Parse(partTicketByPart[1]);
                    //    }
                    //}

                    //if (DefaultTicketResult != null && DefaultTicketResult.Count > 0)
                    //    startTicket = DefaultTicketResult[0].TicketNumber;

                    //var partTicketByPart = startTicket.Split('-');

                    //if (!string.IsNullOrEmpty(Properties.Settings.Default.TicketNumber))
                    //    ticketNumber = int.Parse(Properties.Settings.Default.TicketNumber);
                    //else
                    //{
                    //    if (partTicketByPart.Length == 2)
                    //    {
                    //        ticketNumber = int.Parse(partTicketByPart[1]);
                    //    }
                    //}

                    if (Properties.Settings.Default.TicketExport != null && Properties.Settings.Default.TicketExport > ticketExport)
                        ticketExport = Properties.Settings.Default.TicketExport;

                    // while (ticketNumber <globalTicketNumber)
                    while (ticketExport < DateTime.Today)
                    {
                        if (bgw.CancellationPending)
                        {
                            RepairShoprUtils.LogWriteLineinHTML("Ticket Exporting Process is Stoped or Cancelled by User", MessageSource.Ticket, "", messageType.Warning);
                            bgw.ReportProgress(100, index_);
                            return;
                        }
                        index_ = 1;
                        percentage_ = 1;
                        ticketIndex = 1;
                        //CommitCRM.ObjectQuery<CommitCRM.Ticket> Tickets = new CommitCRM.ObjectQuery<CommitCRM.Ticket>(CommitCRM.LinkEnum.linkAND, 510);
                        //string startIndex=partTicketByPart[0]+"-"+ticketNumber;
                        //Tickets.AddCriteria(CommitCRM.Ticket.Fields.TicketNumber, CommitCRM.OperatorEnum.opGreaterThanOrEqual, startIndex);
                        //string final=partTicketByPart[0]+"-"+(ticketNumber+30);
                        //Tickets.AddCriteria(CommitCRM.Ticket.Fields.TicketNumber, CommitCRM.OperatorEnum.opLessThan, final);

                        RepairShoprUtils.LogWriteLineinHTML(" Loading Tickets from " + ticketExport.ToString() + " to  " + ticketExport.AddMonths(1).ToString(), MessageSource.Ticket, "", messageType.Information);
                        CommitCRM.ObjectQuery<CommitCRM.Ticket> Tickets = new CommitCRM.ObjectQuery<CommitCRM.Ticket>(CommitCRM.LinkEnum.linkAND);
                        Tickets.AddCriteria(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.OperatorEnum.opGreaterThan, ticketExport);
                        Tickets.AddCriteria(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.OperatorEnum.opLessThan, ticketExport.AddMonths(1));
                        Tickets.AddSortExpression(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.SortDirectionEnum.sortASC);
                        _statusMessage = string.Format("Loading Ticket from {0:y}.., it will take 2-3 mintues", ticketExport);
                        bgw.ReportProgress(percentage_, index_);

                        //_statusMessage = string.Format("Loading Ticket from Ticket Number: {0}, it will take 2-3 mintues",startIndex);
                        var CommitCRMTicketLists = new List<CommitCRM.Ticket>();
                        CommitCRMTicketLists = Tickets.FetchObjects();

                        if (CommitCRMTicketLists != null)
                        {
                            ticketCount = CommitCRMTicketLists.Count;
                        }

                        //if (ticketCount == 0) 
                        //{
                        //    break;
                        //}

                        _statusMessage = "Sending to RepairShopr..";
                        bgw.ReportProgress(percentage_, index_);
                        foreach (CommitCRM.Ticket ticket in CommitCRMTicketLists)
                        {
                            try
                            {
                                if (bgw.CancellationPending)
                                {
                                    RepairShoprUtils.LogWriteLineinHTML("Ticket Exporting Process is Stoped or Cancelled by User", MessageSource.Ticket, "", messageType.Warning);
                                    bgw.ReportProgress(100, index_);
                                    break;
                                }

                                string ticketId = string.Empty;
                                using (SQLiteCommand cmdItemAlready = new SQLiteCommand(string.Format("SELECT RTicketId FROM Ticket WHERE ticketId='{0}'", ticket.TicketREC_ID), conn))
                                {
                                    using (SQLiteDataReader reader = cmdItemAlready.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            ticketId = reader[0].ToString();
                                        }
                                    }
                                }

                                if (!string.IsNullOrEmpty(ticketId))
                                {
                                    RepairShoprUtils.LogWriteLineinHTML(string.Format("Ticket With Ticket Number:{0} and Description: {1} and  is already exported", ticket.TicketNumber, ticket.Description), MessageSource.Ticket, "", messageType.Warning);

                                    percentage_ = (100 * index_) / ticketCount;
                                    _statusMessage = string.Format("Ticket : {0} is already Exported so, it is skipping", ticket.TicketNumber);
                                    bgw.ReportProgress(percentage_, index_);
                                    index_++;
                                    //customerIndex++;
                                    continue;
                                }

                                string customerId = string.Empty;
                                using (SQLiteCommand cmdItemAlready = new SQLiteCommand(string.Format("SELECT CustomerId FROM Account WHERE AccountId='{0}'", ticket.AccountREC_ID), conn))
                                {
                                    using (SQLiteDataReader reader = cmdItemAlready.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            customerId = reader[0].ToString();
                                        }
                                    }
                                }

                                if (string.IsNullOrEmpty(customerId))
                                {
                                    customerId = CreateContactForMissingTicket(ticket, conn);
                                    if (string.IsNullOrEmpty(customerId))
                                    {
                                        RepairShoprUtils.LogWriteLineinHTML("Unable to locate Account with Ticket : " + ticket.Description, MessageSource.Ticket, "", messageType.Warning);
                                        percentage_ = (100 * index_) / ticketCount;
                                        _statusMessage = string.Format("Unable to locate customer Id for Ticket : {0}, it is skipping", ticket.Description);
                                        bgw.ReportProgress(percentage_, index_);
                                        index_++;
                                        //ticketIndex++;
                                        continue;
                                    }
                                }
                                RepairShoprUtils.LogWriteLineinHTML("Creating ticket with Ticket Number  : " + ticket.TicketNumber, MessageSource.Ticket, "", messageType.Information);

                                RepairShoprUtils.LogWriteLineinHTML(string.Format("Ticket has following Information :  Subject : {0}, Customer Id: {1},Problem Type :{2},comment_subject:{3}", ticket.Description, customerId, ticket.TicketType, ticket.Status_Text), MessageSource.Ticket, "", messageType.Information);

                                NameValueCollection myNameValueCollection = new NameValueCollection();
                                string subject = string.Empty;
                                if (!string.IsNullOrEmpty(ticket.Description))
                                {
                                    if (ticket.Description.Length > 255)
                                    {
                                        subject = ticket.Description.Substring(0, 255);
                                        RepairShoprUtils.LogWriteLineinHTML(string.Format("Ticket subject is truncate to length 255, Origional Ticket description is {0} and after Truncation is {1}", ticket.Description, subject), MessageSource.Ticket, "", messageType.Warning);
                                    }
                                    else
                                        subject = ticket.Description;
                                }

                                if (_defaultLocationId.HasValue)
                                    myNameValueCollection.Add("location_id", _defaultLocationId.Value.ToString());
                                myNameValueCollection.Add("subject", subject);
                                myNameValueCollection.Add("customer_id", customerId);
                                if (!string.IsNullOrEmpty(ticket.TicketType))
                                    myNameValueCollection.Add("problem_type", ticket.TicketType);
                                myNameValueCollection.Add("status", "Resolved");
                                myNameValueCollection.Add("comment_body", GetCommentValue(ticket));
                                if (!string.IsNullOrEmpty(ticket.Status_Text))
                                    myNameValueCollection.Add("comment_subject", ticket.Status_Text);
                                myNameValueCollection.Add("comment_hidden", "1");
                                myNameValueCollection.Add("comment_do_not_email", "1");
                                string createDate = HttpUtility.UrlEncode(ticket.UpdateDate.ToString("yyyy-MM-dd H:mm:ss"));
                                myNameValueCollection.Add("created_at", createDate);

                                var newTicket = RepairShoprUtils.ExportTicket(myNameValueCollection);
                                if (newTicket != null)
                                {
                                    percentage_ = (100 * index_) / ticketCount;
                                    _statusMessage = string.Format("Exported ( {0}/{1} ) of Ticket", ticketIndex, ticketCount);
                                    bgw.ReportProgress(percentage_, index_);
                                    using (
                                        SQLiteCommand cmdINewItem =
                                            new SQLiteCommand(
                                                string.Format(
                                                    "INSERT INTO  Ticket (TicketId,RTicketId) VALUES('{0}','{1}')",
                                                    ticket.TicketREC_ID, newTicket.Id), conn))
                                    {
                                        cmdINewItem.ExecuteNonQuery();
                                    }

                                    RepairShoprUtils.LogWriteLineinHTML("Successfully Exported New Ticket in RepairShopr ", MessageSource.Ticket, "", messageType.Information);
                                    ticketIndex++;
                                }
                                else
                                {
                                    RepairShoprUtils.LogWriteLineinHTML("Failed Exported New Ticket in RepairShopr " + ticket.Description, MessageSource.Ticket, "", messageType.Warning);
                                }
                            }
                            catch (Exception ex)
                            {
                                RepairShoprUtils.LogWriteLineinHTML("Failed to Export New Ticket. Due to " + ex.Message, MessageSource.Ticket, ex.StackTrace, messageType.Error);
                            }
                            index_++;
                        }

                        RepairShoprUtils.LogWriteLineinHTML("Sucessfull Loaded Ticket up to   " + ticketExport.AddMonths(1).ToString(), MessageSource.Ticket, "", messageType.Information);
                        Properties.Settings.Default.TicketExport = ticketExport;
                        Properties.Settings.Default.Save();
                        ticketExport = ticketExport.AddMonths(1);
                    }
                    isCompleteTicket = true;
                    bgw.ReportProgress(100, index_);
                }
                #endregion
            }
        }

        public string CreateContactForMissingTicket(CommitCRM.Ticket ticket, SQLiteConnection conn)
        {
            string customerID = string.Empty;
            CommitCRM.ObjectQuery<CommitCRM.Account> AccountSearch = new CommitCRM.ObjectQuery<CommitCRM.Account>();
            AccountSearch.AddCriteria(CommitCRM.Account.Fields.AccountREC_ID, CommitCRM.OperatorEnum.opEqual, ticket.AccountREC_ID);
            List<CommitCRM.Account> Accounts = AccountSearch.FetchObjects();
            foreach (CommitCRM.Account account in Accounts)
            {
                try
                {
                    RepairShoprUtils.LogWriteLineinHTML(string.Format("Creating Account with last Name : {0}", account.LastName), MessageSource.Customer, "", messageType.Information);
                    NameValueCollection myNameValueCollection = new NameValueCollection();
                    string fullname = account.GetFieldValue("FLDCRDCONTACT");
                    myNameValueCollection.Add("business_name", account.CompanyName);
                    if (!string.IsNullOrEmpty(fullname) && !string.IsNullOrEmpty(account.LastName))
                        myNameValueCollection.Add("firstname", fullname.Replace(account.LastName, string.Empty));
                    else
                        myNameValueCollection.Add("firstname", fullname);
                    myNameValueCollection.Add("lastname", account.LastName);
                    if (account.EmailAddress1.Contains("@"))
                        myNameValueCollection.Add("email", account.EmailAddress1);
                    myNameValueCollection.Add("phone", account.Phone1);
                    myNameValueCollection.Add("mobile", account.Phone2);
                    myNameValueCollection.Add("address", account.AddressLine1);
                    myNameValueCollection.Add("address_2", account.AddressLine2);
                    myNameValueCollection.Add("city", account.City);
                    myNameValueCollection.Add("state", account.State);
                    myNameValueCollection.Add("zip", account.Zip);
                    myNameValueCollection.Add("notes", account.Notes);
                    var newCustomer = RepairShoprUtils.ExportCustomer(myNameValueCollection);
                    if (newCustomer != null)
                    {
                        using (SQLiteCommand cmdINewItem = new SQLiteCommand(string.Format("INSERT INTO  Account (AccountId,CustomerId) VALUES('{0}','{1}')", account.AccountREC_ID, newCustomer.Id), conn))
                            cmdINewItem.ExecuteNonQuery();
                        customerID = newCustomer.Id;
                        break;
                    }
                    else
                    {
                        RepairShoprUtils.LogWriteLineinHTML(string.Format("Faile to create account : {0} for Ticket : {1}", account.LastName, ticket.Description), MessageSource.Customer, "", messageType.Warning);
                    }
                }
                catch (Exception ex)
                {
                    RepairShoprUtils.LogWriteLineinHTML(string.Format("Faile to create account : {0} for Ticket : {1} dues to {3}", account.LastName, ticket.Description, ex.Message), MessageSource.Customer, ex.StackTrace, messageType.Error);
                }
            }
            return customerID;
        }

        private string GetCommentValue(CommitCRM.Ticket ticket)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("Ticket Number: {0}", ticket.TicketNumber.Replace("-", string.Empty)));
            try
            {
                sb.AppendFormat(" , Cause: {0}", ticket.GetFieldValue("FLDTKTCAUSE"));
                sb.AppendFormat(" , Source: {0}", ticket.Source);
                sb.AppendFormat(" , Category: {0}", ticket.GetFieldValue("FLDTKTCATEGORY"));
                sb.AppendFormat(" , Note: {0}", ticket.Notes);
                sb.AppendFormat(" , Resolution : {0}", ticket.Resolution);

                CommitCRM.ObjectQuery<CommitCRM.Charge> ChargeSearch = new CommitCRM.ObjectQuery<CommitCRM.Charge>();
                ChargeSearch.AddCriteria(CommitCRM.Charge.Fields.TicketREC_ID, CommitCRM.OperatorEnum.opEqual, ticket.TicketREC_ID);
                List<CommitCRM.Charge> Charges = ChargeSearch.FetchObjects();
                foreach (CommitCRM.Charge Charge in Charges)
                {
                    sb.AppendFormat(", Charge :{0}", Charge.Description);
                    sb.AppendFormat(", Amount : {0}", Charge.GetFieldValue("FLDSLPBILLTOTAL"));
                    sb.AppendFormat(", Quantity: {0}", Charge.GetFieldValue("FLDSLPQUANTITY"));
                    sb.AppendFormat(", Date : {0}", Charge.Date);
                }

                CommitCRM.ObjectQuery<CommitCRM.HistoryNote> HistoryNoteSearch = new CommitCRM.ObjectQuery<CommitCRM.HistoryNote>();
                HistoryNoteSearch.AddCriteria(CommitCRM.HistoryNote.Fields.RelLinkREC_ID, CommitCRM.OperatorEnum.opEqual, ticket.TicketREC_ID);
                List<CommitCRM.HistoryNote> HistoryNotes = HistoryNoteSearch.FetchObjects();
                foreach (CommitCRM.HistoryNote HistoryNote in HistoryNotes)
                {
                    sb.AppendFormat(" History Note :{0}", HistoryNote.Description);
                    sb.AppendFormat(", Date : {0}", HistoryNote.Date);
                }
            }
            catch (Exception ex)
            {
                RepairShoprUtils.LogWriteLineinHTML(string.Format("Failed to get Charge, HistoryNote of Ticket : {0} . Due to {1}", ticket.Description, ex.Message), MessageSource.Ticket, ex.StackTrace, messageType.Warning);
            }
            return sb.ToString();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage < 100)
            {
                progressBar1.Value = e.ProgressPercentage;
                buttonExport.Enabled = false;
                label3.Text = _statusMessage;
            }
            else
            {
                progressBar1.Value = e.ProgressPercentage;
                if (_exportCustomer && _exportTicket)
                {
                    if (isCompleteCustomer && isCompleteTicket)
                    {
                        buttonStop.Enabled = false;
                        buttonExport.Enabled = true;
                        label3.Text = "Exporting Process is Completed";
                        RepairShoprUtils.LogWriteLineinHTML("Exporting Process is Completed ", MessageSource.Complete, "", messageType.Information);
                    }
                }
                else if (_exportCustomer)
                {
                    if (isCompleteCustomer)
                    {
                        buttonStop.Enabled = false;
                        buttonExport.Enabled = true;
                        label3.Text = "Exporting Process is Completed";
                        RepairShoprUtils.LogWriteLineinHTML("Exporting Process is Completed ", MessageSource.Complete, "", messageType.Information);
                    }
                }
                else if (_exportTicket)
                {
                    if (isCompleteTicket)
                    {
                        buttonStop.Enabled = false;
                        buttonExport.Enabled = true;
                        label3.Text = "Exporting Process is Completed";
                        RepairShoprUtils.LogWriteLineinHTML("Exporting Process is Completed ", MessageSource.Complete, "", messageType.Information);
                    }
                }
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void viewLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(RepairShoprUtils.path);
        }

        private void textBoxUserName_TextChanged(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            label1.Text = string.Empty;
        }

        private void textBoxPassWord_TextChanged(object sender, EventArgs e)
        {
            errorProvider2.Clear();
            label1.Text = string.Empty;
        }

        private void checkBoxExportCustomer_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxExportTicket.Enabled = checkBoxExportCustomer.Checked;
            checkBoxExportTicket.Checked = checkBoxExportCustomer.Checked;
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            if (bgw.IsBusy)
            {
                bgw.CancelAsync();
                progressBar1.Value = 100;
                label3.Text = "Exporting process is cancelled";
                buttonStop.Enabled = false;
            }
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new AboutBox1().ShowDialog();
        }

        private void resetConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int result = -1;
            int rsult = -1;
            RepairShoprUtils.LogWriteLineinHTML("Reseting Configuration including Database.. ", MessageSource.Initialization, "", messageType.Information);
            using (SQLiteConnection conn = new SQLiteConnection("data source=" + _path + ";PRAGMA journal_mode=WAL;Password=shyam;"))
            {
                conn.Open();
                using (SQLiteCommand cmdAccountDelete = new SQLiteCommand("DELETE From Account", conn))
                {
                    result = cmdAccountDelete.ExecuteNonQuery();
                }
                using (SQLiteCommand cmdTicketDelete = new SQLiteCommand("DELETE From Ticket", conn))
                {
                    rsult = cmdTicketDelete.ExecuteNonQuery();
                }

                CommitCRM.ObjectQuery<CommitCRM.Ticket> DefaultTickets = new CommitCRM.ObjectQuery<CommitCRM.Ticket>(CommitCRM.LinkEnum.linkAND, 1);
                DefaultTickets.AddSortExpression(CommitCRM.Ticket.Fields.UpdateDate, CommitCRM.SortDirectionEnum.sortASC);
                List<CommitCRM.Ticket> DefaultTicketResult = DefaultTickets.FetchObjects();
                string ticketNumber = string.Empty;
                if (DefaultTicketResult != null && DefaultTicketResult.Count > 0)
                    ticketNumber = DefaultTicketResult[0].TicketNumber;
                Properties.Settings.Default.TicketNumber = ticketNumber;

                CommitCRM.ObjectQuery<CommitCRM.Account> DefaultAccounts = new CommitCRM.ObjectQuery<CommitCRM.Account>(CommitCRM.LinkEnum.linkAND, 1);
                DefaultAccounts.AddSortExpression(CommitCRM.Account.Fields.CreationDate, CommitCRM.SortDirectionEnum.sortASC);
                List<CommitCRM.Account> DefaultAccountResult = DefaultAccounts.FetchObjects();
                DateTime customerExport = Directory.GetCreationTime(installedLocation);
                if (DefaultAccountResult != null && DefaultAccountResult.Count > 0)
                    customerExport = DefaultAccountResult[0].CreationDate;
                Properties.Settings.Default.CustomerExport = customerExport;
                Properties.Settings.Default.TicketExport = customerExport;
                Properties.Settings.Default.Save();

                if (result != -1 && rsult != -1)
                {
                    RepairShoprUtils.LogWriteLineinHTML("Reset Configuration including Database successfull.. ", MessageSource.Initialization, "", messageType.Information);

                    MessageBox.Show("Configuration is reset successfull", "RepairShoprApps", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void supportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("http://feedback.repairshopr.com/knowledgebase/articles/796074");
        }
    }
}