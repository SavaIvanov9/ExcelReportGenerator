using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ClientsRepeatability.Data;

namespace ClientsRepeatability.ExcelReportGenerator.DataProcessing
{
    public class DataProcessor
    {
        private readonly List<V_Clients> _clients;
        private readonly List<V_Loans> _loans;
        private readonly List<Liability> _liabilities;
        private readonly string _operation;
        private readonly DateTime _dateToday;
        private readonly int _startIndex;
        private readonly int _elementsToProcessCount;
        private readonly object _additionalData;
        private object _generatedData;

        public DataProcessor(List<V_Clients> clients, List<V_Loans> loans,
            List<Liability> liabilities, DateTime dateToday, string operation,
            int startIndex, int elementsToProcessCount,
            object additionalData = null)
        {
            this._clients = clients;
            this._loans = loans;
            this._liabilities = liabilities;
            this._dateToday = dateToday;
            this._operation = operation;
            this._startIndex = startIndex;
            this._elementsToProcessCount = elementsToProcessCount;
            this._additionalData = additionalData;
            //this._additionamessage = additionalMessage;
        }

        public object Data
        {
            get { return this._generatedData; }
        }

        public void ProcessData()
        {
            if (_operation == "clients")
            {
                ProcessClients();
            }
            else if (_operation == "loans")
            {
                ProcessLoans();
            }
            else if (_operation == "findFirstLoan")
            {
                ProcessFirstLoan();
            }
        }

        private void ProcessClients()
        {
            var lastIndex = this._startIndex + this._elementsToProcessCount;
            var result = new List<Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>>();

            for (int i = _startIndex; i < lastIndex; i++)
            {
                var client = _clients[i];

                var firstLoan = _loans.FirstOrDefault(
                        loan => loan.PINOrCRN == client.PINOrCRN &&
                        _liabilities.Exists(liability => liability.ID_Owner == loan.RID));

                if (firstLoan != null)
                {
                    //-----------------

                    var dataManager1 = new DataManager();
                    var filteredLoans = (List<V_Loans>)
                        dataManager1.GenerateData(_clients, _loans, _liabilities, _dateToday, "findFirstLoan",
                        client);

                    //---------------------

                    //find first loan
                    foreach (var loan in filteredLoans)
                    {
                        if (client.PINOrCRN == loan.PINOrCRN &&
                            _liabilities.Exists(x => x.ID_Owner == loan.RID))
                        {
                            var date = _liabilities
                                .First(liability => liability.ID_Owner == loan.RID)
                                .CreationDate;

                            var currentFirstLoanDate = _liabilities
                                .First(liability => liability.ID_Owner == firstLoan.RID)
                                .CreationDate;

                            if (date < currentFirstLoanDate)
                            {
                                firstLoan = loan;
                            }
                        }
                    }

                    //check is first loan in the interval
                    var firstLoanDate = _liabilities
                        .First(liability => liability.ID_Owner == firstLoan.RID)
                        .CreationDate;

                    var clientLoans = new List<Tuple<V_Loans, DateTime>>();

                    if (firstLoanDate >= _dateToday.AddMonths(-30) &&
                        firstLoanDate < _dateToday.AddMonths(-6))
                    {
                        ////---
                        var dataManager2 = new DataManager();
                        clientLoans = (List<Tuple<V_Loans, DateTime>>) dataManager2
                            .GenerateData(_clients, _loans, _liabilities, _dateToday,
                                "loans", new Tuple<V_Clients, DateTime?>(client, firstLoanDate));
                        ////---

                        ////find other loans
                        //foreach (var loan in _loans)
                        //{
                        //    if (client.PINOrCRN == loan.PINOrCRN &&
                        //        _liabilities.Exists(x => x.ID_Owner == loan.RID))
                        //    {
                        //        if (_liabilities.First(l => l.ID_Owner == loan.RID).CreationDate > firstLoanDate)
                        //        {
                        //            clientLoans.Add(new Tuple<V_Loans, DateTime>(loan,
                        //                (DateTime)_liabilities.First(l => l.ID_Owner == loan.RID).CreationDate));
                        //        }
                        //    }
                        //}
                    }

                    if (clientLoans.Count != 0)
                    {
                        result.Add(
                            new Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>(
                                client, firstLoan.LoanNumber.ToString(), clientLoans, (DateTime)firstLoanDate));
                    }
                }

                _generatedData = result;
            }
        }

        private void ProcessLoans()
        {
            var lastIndex = this._startIndex + this._elementsToProcessCount;
            var clientLoans = new List<Tuple<V_Loans, DateTime>>();

            for (int i = _startIndex; i < lastIndex; i++)
            {
                var loan = _loans[i];

                var tuple = (Tuple<V_Clients, DateTime?>)_additionalData;

                if (tuple.Item1.PINOrCRN == loan.PINOrCRN &&
                    _liabilities.Exists(x => x.ID_Owner == loan.RID))
                {
                    if (_liabilities.First(l => l.ID_Owner == loan.RID).CreationDate > tuple.Item2)
                    {
                        clientLoans.Add(new Tuple<V_Loans, DateTime>(loan,
                            (DateTime) _liabilities.First(l => l.ID_Owner == loan.RID).CreationDate));
                    }
                }
            }

            _generatedData = clientLoans;
        }

        private void ProcessFirstLoan()
        {
            var lastIndex = this._startIndex + this._elementsToProcessCount;
            var firstLoan = new V_Loans();

            for (int i = _startIndex; i < lastIndex; i++)
            {
                var loan = _loans[i];
                var client = (V_Clients) _additionalData;
                if (client.PINOrCRN == loan.PINOrCRN &&
                    _liabilities.Exists(x => x.ID_Owner == loan.RID))
                {
                    var date = _liabilities
                        .First(liability => liability.ID_Owner == loan.RID)
                        .CreationDate;

                    var firstOrDefault = _liabilities
                        .FirstOrDefault(liability => liability.ID_Owner == firstLoan.RID);
                    if (firstOrDefault != null)
                    {
                        var currentFirstLoanDate = firstOrDefault
                            .CreationDate;

                        if (date < currentFirstLoanDate)
                        {
                            firstLoan = loan;
                        }
                    }
                }
            }

            _generatedData = firstLoan;
        }
    }
}
