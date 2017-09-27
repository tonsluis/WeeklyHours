using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WeeklyHoursXlReportBuilder
{
  /// <summary>
  /// 
  /// </summary>
  public class Worker
  {
    /// <summary>
    /// Personell Number.
    /// </summary>
    public int Number { get; set; }

    /// <summary>
    /// Name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// ProjectNumber
    /// </summary>
    public string ProjectNumber { get; set; }

    /// <summary>
    /// Year
    /// </summary>
    public int Year { get; set; }

    /// <summary>
    /// Week
    /// </summary>
    public int Week { get; set; }

    /// <summary>
    /// Hours
    /// </summary>
    public Double Hours { get; set; }

  }
}
