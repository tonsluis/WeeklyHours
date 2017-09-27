

namespace WeeklyHoursXlReportBuilder
{
  #region Directives
  using System;
  using System.Collections.Generic;
  using Database;


  #endregion
  public class Host
  {

    public Host()
    {

    }


    public List<Worker> GetReportData()
    {
      List<Worker> data = new List<Worker>(); 
      using (var repo = new DailyReportRepository())
      {
        data =  repo.GetHoursReport();
      }
      return data;
    }

  }
}
