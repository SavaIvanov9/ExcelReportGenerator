using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using ClientsRepeatability.Data;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace ClientsRepeatability.ExcelReportGenerator.DataProcessing
{
    public class DataManager
    {
        private readonly int _coresCount;
        private readonly List<Thread> _threads;
        private readonly List<DataProcessor> _dataProcessors;

        public DataManager()
        {
            this._coresCount = Environment.ProcessorCount;
            this._threads = new List<Thread>(_coresCount);
            this._dataProcessors = new List<DataProcessor>(_coresCount);
        }

        public object GenerateData(List<V_Clients> clients, List<V_Loans> loans,
            List<Liability> liabilities, DateTime dateToday, string operation,
            object additionalData = null)
        {
            if (operation == "clients")
            {
                var elementsPerCore = clients.Count / _coresCount;
                var elementsLeftOver = clients.Count % _coresCount;

                for (int i = 0; i < _coresCount; i++)
                {
                    var startIndex = i * elementsPerCore;
                    var elementsToProcessCount = elementsPerCore;

                    if (i == _coresCount - 1)
                    {
                        elementsToProcessCount += elementsLeftOver;
                    }

                    var seedProcessor = new DataProcessor(
                        clients, loans, liabilities, dateToday,
                        operation, startIndex, elementsToProcessCount,
                        additionalData);
                    _dataProcessors.Add(seedProcessor);

                    var thread = new Thread(seedProcessor.ProcessData);
                    _threads.Add(thread);
                    thread.Start();
                }

                var dataToReturn = new List<Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>>();

                for (int i = 0; i < _threads.Count; i++)
                {
                    _threads[i].Join();
                    dataToReturn.AddRange(
                        _dataProcessors[i].Data as
                            IEnumerable<Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>>);
                }

                return dataToReturn;
            }
            else if (operation == "loans")
            {
                var elementsPerCore = loans.Count / _coresCount;
                var elementsLeftOver = loans.Count % _coresCount;

                for (int i = 0; i < _coresCount; i++)
                {
                    var startIndex = i * elementsPerCore;
                    var elementsToProcessCount = elementsPerCore;

                    if (i == _coresCount - 1)
                    {
                        elementsToProcessCount += elementsLeftOver;
                    }

                    var seedProcessor = new DataProcessor(
                        clients, loans, liabilities, dateToday,
                        operation, startIndex, elementsToProcessCount,
                        additionalData);
                    _dataProcessors.Add(seedProcessor);

                    var thread = new Thread(seedProcessor.ProcessData);
                    _threads.Add(thread);
                    thread.Start();
                }

                var dataToReturn = new List<Tuple<V_Loans, DateTime>>();

                for (int i = 0; i < _threads.Count; i++)
                {
                    _threads[i].Join();
                    dataToReturn.AddRange(
                        _dataProcessors[i].Data as IEnumerable<Tuple<V_Loans, DateTime>>);
                }

                return dataToReturn;
            }
            else if (operation == "findFirstLoan")
            {
                var elementsPerCore = loans.Count / _coresCount;
                var elementsLeftOver = loans.Count % _coresCount;

                for (int i = 0; i < _coresCount; i++)
                {
                    var startIndex = i * elementsPerCore;
                    var elementsToProcessCount = elementsPerCore;

                    if (i == _coresCount - 1)
                    {
                        elementsToProcessCount += elementsLeftOver;
                    }

                    var seedProcessor = new DataProcessor(
                        clients, loans, liabilities, dateToday,
                        operation, startIndex, elementsToProcessCount,
                        additionalData);
                    _dataProcessors.Add(seedProcessor);

                    var thread = new Thread(seedProcessor.ProcessData);
                    _threads.Add(thread);
                    thread.Start();
                }

                var dataToReturn = new List<V_Loans>();

                for (int i = 0; i < _threads.Count; i++)
                {
                    _threads[i].Join();
                    dataToReturn.Add(_dataProcessors[i].Data as V_Loans);
                }

                return dataToReturn;
            }

            return null;
        }
    }
}
