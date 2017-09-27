

namespace WeeklyHoursXlReportBuilder
{

  #region Directives
  // Standard .NET Directives
  using System;
  using System.IO;
  using System.Collections.Generic;
  using Microsoft.Office.Interop.Excel;
  // CSi Directives
  using Database;

  #endregion



  class Program
  {
    static void Main(string[] args)
    {
      int wk = Helper.IsoWeek(DateTime.Now);

      string strFilePath = $"C:\\Temp\\MSS_WeeklyHourReport_{DateTime.Now.Year:D4}wk{wk:D2}.xlsx";
      if (File.Exists(strFilePath))
      {
        File.Delete(strFilePath);
      }

      ADOFactory.SetCreateMethod("standard", DefaultConnection, true);




      // Application
      Application xlApp = new Application
      {
        DisplayAlerts = true
      };

      // Workbook
      Workbook xlWorkbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
      //Worksheet
      Worksheet xlWorkSheet = (Worksheet)xlWorkbook.Worksheets[1];
      //xlWorkSheet.Name = "Uren wk37";
      // Set the App to visible.
      xlApp.Visible = false;

      xlWorkSheet.Columns[HourlyReportColumnNames.PersonnelNumber].ColumnWidth = 18;
      xlWorkSheet.Columns[HourlyReportColumnNames.Name].ColumnWidth = 25;
      xlWorkSheet.Columns[HourlyReportColumnNames.Project].ColumnWidth = 8;
      xlWorkSheet.Columns[HourlyReportColumnNames.Year].ColumnWidth = 6;
      xlWorkSheet.Columns[HourlyReportColumnNames.Week].ColumnWidth = 6;
      xlWorkSheet.Columns[HourlyReportColumnNames.Hours].ColumnWidth = 6;

      xlWorkSheet.Name = $"Uren wk{wk:D2}";

      Host host = new Host();

      Console.WriteLine("Getting data from Host....");

      List<Worker> data = host.GetReportData();

      xlWorkSheet.Cells[1, HourlyReportColumnNames.PersonnelNumber] = "Personnel number";
      xlWorkSheet.Cells[1, HourlyReportColumnNames.Name] = "Name";
      xlWorkSheet.Cells[1, HourlyReportColumnNames.Project] = "Project";
      xlWorkSheet.Cells[1, HourlyReportColumnNames.Year] = "Year";
      xlWorkSheet.Cells[1, HourlyReportColumnNames.Week] = "Week";
      xlWorkSheet.Cells[1, HourlyReportColumnNames.Hours] = "Hours";

      // Font to bold for the titles.
      var startCell = (Range)xlWorkSheet.Cells[1, HourlyReportColumnNames.PersonnelNumber];
      var endCell = (Range)xlWorkSheet.Cells[1, HourlyReportColumnNames.Hours];
      var _range = xlWorkSheet.Range[startCell, endCell];

      _range.Font.Bold = true;
      _range.Font.Size = 12;


      // Allignment of the columns.
      xlWorkSheet.Range["A:A", "A:A"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
      xlWorkSheet.Range["C:C", "F:F"].HorizontalAlignment = XlHAlign.xlHAlignCenter;

      // Lock Title row.
      Range _lockCell = (Range)xlWorkSheet.Cells[2, 1];
      _lockCell.Activate();
      _lockCell.Application.ActiveWindow.FreezePanes = true;

      // Populate the report.
      int rowCount = 2;
      int maxLines = data.Count;
      foreach (var item in data)
      {
        xlWorkSheet.Cells[rowCount, HourlyReportColumnNames.PersonnelNumber] = item.Number;
        xlWorkSheet.Cells[rowCount, HourlyReportColumnNames.Name] = item.Name;
        xlWorkSheet.Cells[rowCount, HourlyReportColumnNames.Project] = item.ProjectNumber;
        xlWorkSheet.Cells[rowCount, HourlyReportColumnNames.Year] = item.Year;
        xlWorkSheet.Cells[rowCount, HourlyReportColumnNames.Week] = item.Week;
        xlWorkSheet.Cells[rowCount, HourlyReportColumnNames.Hours] = item.Hours;
        Console.WriteLine($"Report line {rowCount} - {maxLines}");
        rowCount++;
      }

      // Save it.
      xlWorkbook.SaveAs(strFilePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                  false, false, XlSaveAsAccessMode.xlNoChange,
                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
      xlWorkbook.Close();
      Console.WriteLine($"Done generating file '{strFilePath}'");



      xlApp.Quit();

      System.Diagnostics.Process.Start("excel.exe", strFilePath);

    }

    private static void LoadData()
    {
    }


    private static IADOConnection DefaultConnection()
    {
      string connString = string.Format(@"Data Source =DBSNL; Initial Catalog = {0}; Integrated Security = True; ", "master");
      return ADOConnection.Create(connString);
    }




  }
}
