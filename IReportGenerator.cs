namespace ClientsRepeatability.ExcelReportGenerator
{
    using Data;
    using System;
    using System.Collections.Generic;

    public interface IReportGenerator
    {
        void GenerateReport(DateTime date);
    }
}
