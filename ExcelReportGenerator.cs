using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClientsRepeatability.Data;
using ClientsRepeatability.Data.Fakes;
using ClientsRepeatability.ExcelReportGenerator.DataProcessing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Telerik.JustMock;

namespace ClientsRepeatability.ExcelReportGenerator
{
    public class ExcelReportGenerator : IReportGenerator
    {
        private int _fakeDataCount = 50000;
        public void GenerateReport(DateTime date)
        {
            var reportsFolderName = "Reports";
            CheckDirectory(reportsFolderName);

            string currentFile = $"{DateTime.Now:yyyy-MM-dd_hh;mm;ss}.xlsx";
            string path = Path.Combine(
                Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location),
                    $@"{reportsFolderName}\{currentFile}");

            var reportData = PrepairReport(date);
            CreateFile(path, reportData);
        }

        private void CreateFile(string path, MemoryStream data)
        {
            if (!File.Exists(path))
            {
                var fileInfo = new FileInfo(path);

                using (var p = new ExcelPackage(fileInfo))
                {
                    p.Load(data);
                    p.Save();
                }
                //using (StreamWriter file = File.CreateText(path))
                //{
                //    data.WriteTo(file);
                //    file.Write(data);
                //}
            }
        }

        private void CheckDirectory(string folderName)
        {
            string dirName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string pathString = Path.Combine(dirName, folderName);
            if (!Directory.Exists(pathString))
            {
                Directory.CreateDirectory(pathString);
            }
        }

        private ICompanyData CreateFakeData(int count)
        {
            var fakeData = Mock.Create<ICompanyData>();
            var fakeClients = new List<V_Clients>();
            var fakeLoans = new List<V_Loans>();
            var fakeLiabilities = new List<Liability>();

            for (int i = 0; i < count; i++)
            {
                fakeClients.Add(new V_Clients
                {
                    ClientType = 0,
                    PINOrCRN = "pin-" + i,
                });

                fakeLoans.Add(new V_Loans
                {
                    RID = i,
                    PINOrCRN = "pin-" + i,
                    LoanNumber = 1000 + i
                });
                fakeLoans.Add(new V_Loans
                {
                    RID = i + 2000,
                    PINOrCRN = "pin-" + i,
                    LoanNumber = 1000 + i
                });
                fakeLoans.Add(new V_Loans
                {
                    RID = i + 3000,
                    PINOrCRN = "pin-" + i,
                    LoanNumber = 1000 + i
                });

                fakeLiabilities.Add(new Liability
                {
                    ID_CL_Liability = -1,
                    ID_OrigLiability = null,
                    ID_Owner = i,
                    CreationDate = DateTime.Now.AddMonths(-12)
                });
                fakeLiabilities.Add(new Liability
                {
                    ID_CL_Liability = -1,
                    ID_OrigLiability = null,
                    ID_Owner = i + 2000,
                    CreationDate = DateTime.Now.AddMonths(-10)
                });
                fakeLiabilities.Add(new Liability
                {
                    ID_CL_Liability = -1,
                    ID_OrigLiability = null,
                    ID_Owner = i + 3000,
                    CreationDate = DateTime.Now.AddMonths(-8)
                });
            }

            Mock.Arrange(() => fakeData.Clients.All())
                .Returns(() => fakeClients.AsQueryable());
            Mock.Arrange(() => fakeData.Loans.All())
                .Returns(() => fakeLoans.AsQueryable());
            Mock.Arrange(() => fakeData.Liabilities.All())
                .Returns(() => fakeLiabilities.AsQueryable());

            return fakeData;
        }

        private MemoryStream PrepairReport(DateTime date)
        {
            Console.WriteLine("Prepairing report...");

            MemoryStream outputStream = new MemoryStream();

            using (ExcelPackage pcg = new ExcelPackage(outputStream))
            {
                ExcelWorksheet ws = pcg.Workbook.Worksheets.Add("ClientsRepeatability");

                //Header
                ws.Cells[1, 1].Value = "ЕГН";
                ws.Cells[1, 2].Value = "Номер на договора";
                for (int i = 1; i <= 30; i++)
                {
                    ws.Cells[1, i + 2].Value = i;
                }
                ws.Cells[1, 33].Value = "Общо";
                ws.Cells["A1:AG1"].Style.Font.Bold = true;
                ws.Cells["A1:AG1"].Style.Font.Size = 12;
                ws.Cells["A1:AG"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A1:AG"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //Data
                var reportData = ExtractData(date);
                var allTotalLoansPerColumn = new int[30];

                Console.WriteLine("Filling data in report...");

                int row = 2;
                foreach (var client in reportData)
                {
                    ws.Cells[row, 1].Value = client.Item1.PINOrCRN;
                    ws.Cells[row, 2].Value = client.Item2; // document number?

                    var totalLoansPerRow = 0;
                    var month = 1;
                    for (int i = 3; i < 33; i++)
                    {
                        var loansPerMonth = CalculateLoansPerMonth(
                            client.Item3,
                            client.Item4.AddMonths(month));

                        if (loansPerMonth > 0)
                        {
                            ws.Cells[row, i].Value = loansPerMonth;
                            allTotalLoansPerColumn[i - 3] += loansPerMonth;
                            totalLoansPerRow += loansPerMonth;
                        }

                        month++;
                    }

                    ws.Cells[row, 33].Value = totalLoansPerRow;
                    row++;
                }

                ws.Cells[row, 1].Value = "Общо";
                ws.Cells[row, 2].Value = reportData.Count;

                for (int i = 3; i < allTotalLoansPerColumn.Length + 3; i++)
                {
                    ws.Cells[row, i].Value = allTotalLoansPerColumn[i - 3];
                }

                ws.Cells.AutoFitColumns();
                pcg.Save();
                outputStream.Position = 0;
                Console.WriteLine("Report prepaired.");
                return outputStream;
            }
        }

        private List<Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>>
            ExtractData(DateTime date)
        {
            Console.WriteLine("Collecting data...");
            var clients = new List<V_Clients>();
            var loans = new List<V_Loans>();
            var liabilities = new List<Liability>();
            var threads = new List<Thread>();

            var clientsExtractingThread = new Thread(() =>
            {
                //using (var data = new CompanyData("test"))
                using (var data = CreateFakeData(_fakeDataCount))
                {
                    clients = data
                    .Clients
                    .All()
                    .Where(x => x.ClientType == 0)
                    .AsNoTracking()
                    .ToList();
                }
            });
            threads.Add(clientsExtractingThread);

            var loansExtractingThread = new Thread(() =>
            {
                //using (var data = new CompanyData("test"))
                using (var data = CreateFakeData(_fakeDataCount))
                {
                    loans = data
                    .Loans
                    .All()
                    .AsNoTracking()
                    .ToList();
                }
            });
            threads.Add(loansExtractingThread);

            var liabilitiesExtractingThread = new Thread(() =>
            {
                //using (var data = new CompanyData("test"))
                using (var data = CreateFakeData(_fakeDataCount))
                {
                    liabilities = data
                    .Liabilities
                    .All()
                    .Where(
                        x => x.ID_CL_Liability == -1 &&
                        x.ID_OrigLiability == null)
                    .AsNoTracking()
                    .ToList();
                }
            });
            threads.Add(liabilitiesExtractingThread);

            Parallel.ForEach(threads, thread => thread.Start());
            Parallel.ForEach(threads, thread => thread.Join());

            Console.WriteLine("Data collected.");

            var newClients = FilterClients(clients, loans, liabilities, date);
            return newClients;
        }
        
        private List<Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>> FilterClients(
            List<V_Clients> clients,
            List<V_Loans> loans,
            List<Liability> liabilities,
            DateTime dateToday)
        {
            Console.WriteLine("Filtering data...");

            //------------Multithreaded filtering---------------
            //
            var dataManager = new DataManager();
            var result = (List<Tuple<V_Clients, string, List<Tuple<V_Loans, DateTime>>, DateTime>>)
                dataManager.GenerateData(clients, loans, liabilities, dateToday, "clients");

            Console.WriteLine("Data filered.");

            return result;

            ////----------Singlethreaded filtering-------------
            //// 
            //var result = new List<Tuple<V_Clients, string, Dictionary<V_Loans, DateTime>, DateTime>>();

            //foreach (var client in clients)
            //{
            //    //Console.WriteLine("+");
            //    var clientLoans = new Dictionary<V_Loans, DateTime>();
            //    var firstLoan = loans
            //        .FirstOrDefault(loan => loan.PINOrCRN == client.PINOrCRN &&
            //            liabilities.Exists(liability => liability.ID_Owner == loan.RID));

            //    if (firstLoan != null)
            //    {
            //        //find first loan
            //        foreach (var loan in loans)
            //        {
            //            if (client.PINOrCRN == loan.PINOrCRN &&
            //                liabilities.Exists(x => x.ID_Owner == loan.RID))
            //            {
            //                var date = liabilities
            //                    .First(liability => liability.ID_Owner == loan.RID)
            //                    .CreationDate;

            //                var currentFirstLoanDate = liabilities
            //                    .First(liability => liability.ID_Owner == firstLoan.RID)
            //                    .CreationDate;

            //                if (date < currentFirstLoanDate)
            //                {
            //                    firstLoan = loan;
            //                }
            //            }
            //        }

            //        //check is first loan in the interval
            //        var firstLoanDate = liabilities
            //            .First(liability => liability.ID_Owner == firstLoan.RID)
            //            .CreationDate;

            //        if (firstLoanDate >= dateToday.AddMonths(-30) &&
            //            firstLoanDate < dateToday.AddMonths(-6))
            //        {
            //            //find other loans
            //            foreach (var loan in loans)
            //            {
            //                if (client.PINOrCRN == loan.PINOrCRN &&
            //                    liabilities.Exists(x => x.ID_Owner == loan.RID))
            //                {
            //                    if (liabilities.First(l => l.ID_Owner == loan.RID).CreationDate > firstLoanDate)
            //                    {
            //                        clientLoans.Add(loan,
            //                            (DateTime) liabilities.First(l => l.ID_Owner == loan.RID).CreationDate);
            //                    }
            //                }
            //            }
            //        }

            //        if (clientLoans.Count != 0)
            //        {
            //            result.Add(
            //                new Tuple<V_Clients, string, Dictionary<V_Loans, DateTime>, DateTime>(
            //                    client, firstLoan.LoanNumber.ToString(), clientLoans, (DateTime) firstLoanDate));
            //        }
            //    }
            //}
            //Console.WriteLine("Data filered.");

            //return result;
        }

        private int CalculateLoansPerMonth(List<Tuple<V_Loans, DateTime>> data, DateTime month)
        {
            var result = 0;

            foreach (var loan in data)
            {
                if (loan.Item2.Month == month.Month &&
                    loan.Item2.Year == month.Year)
                {
                    result++;
                }
            }

            return result;
        }
    }
}


