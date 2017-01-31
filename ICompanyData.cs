namespace ClientsRepeatability.Data
{
    using System;
    using Repositories;

    public interface ICompanyData : IDisposable
    {
        ClientsRepository Clients
        {
            get;
        }

        LoansRepository Loans
        {
            get;
        }

        LiabilitiesRepository Liabilities
        {
            get;
        }
    }
}
