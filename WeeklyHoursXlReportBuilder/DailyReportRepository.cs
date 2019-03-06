
namespace WeeklyHoursXlReportBuilder
{
    #region Directives
    using System.Collections.Generic;
    using System;
    using System.Text;


    using System.Data.SqlClient;

    using Database;


    #endregion

    public class DailyReportRepository : IDisposable
    {
        public void Dispose()
        {

        }
        public List<Worker> GetHoursReport(IList<int> ids)
        {
            var id = string.Join(",", ids);
            StringBuilder query = new StringBuilder();

            query.AppendLine("SELECT");
            query.AppendLine("PN.T$EMNO as PersoneelNummer,");

            query.AppendLine("PN.T$NAMA as [Name]");
            query.AppendLine(",[Project]");
            query.AppendLine(",[Jaar]");
            query.AppendLine(",[Week]");
            query.AppendLine(", sum([Aantal]) as hours");
            query.AppendLine("FROM[dbsRapportageAlgemeen].[dbo].[vwCSiUrenspecificatieCSI] as UR");
            query.AppendLine("Inner Join[dbsBaaNTables100].[dbo].[tblTTCCOM001100] PN on(cast(UR.Medewerker as float) = PN.[T$EMNO])");
            query.AppendLine($"Where PN.T$EMNO in ({id})");
            query.AppendLine("group by PN.T$EMNO, PN.T$NAMA, Project, [week], Jaar");
            query.AppendLine("Order by PN.T$EMNO, UR.Jaar, UR.[week], PN.T$NAMA");

            List<Worker> returnValue = new List<Worker>();
            using (IADOConnection conn = ADOFactory.Create() as IADOConnection)
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = query.ToString();
                    SqlDataReader reader = (SqlDataReader)cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        Worker worker = new Worker();

                        worker.Number = Convert.ToInt32(reader["PersoneelNummer"]);
                        worker.Name = reader["Name"].ToString();
                        worker.ProjectNumber = reader["Project"].ToString();
                        worker.Year = Convert.ToInt32(reader["Jaar"]);
                        worker.Week = Convert.ToInt32(reader["Week"]);
                        worker.Hours = Convert.ToDouble(reader["Hours"]);

                        returnValue.Add(worker);

                    }

                    reader.Close();
                    reader = null;

                }
            }
            return returnValue;
        }




        /// <summary>
        /// Returns a list of hours per person, per project, per week 
        /// </summary>
        /// <returns></returns>
        public List<Worker> GetHoursReport()
        {
            StringBuilder query = new StringBuilder();

            query.AppendLine("SELECT");
            query.AppendLine("PN.T$EMNO as PersoneelNummer,");

            query.AppendLine("PN.T$NAMA as [Name]");
            query.AppendLine(",[Project]");
            query.AppendLine(",[Jaar]");
            query.AppendLine(",[Week]");
            query.AppendLine(", sum([Aantal]) as hours");
            query.AppendLine("FROM[dbsRapportageAlgemeen].[dbo].[vwCSiUrenspecificatieCSI] as UR");
            query.AppendLine("Inner Join[dbsBaaNTables100].[dbo].[tblTTCCOM001100] PN on(cast(UR.Medewerker as float) = PN.[T$EMNO])");
            query.AppendLine("Where PN.T$EMNO in (11302, 11631, 11742, 11934, 50924, 60619, 65044, 65052, 11771, 11873,11968, 12065)");
            query.AppendLine("group by PN.T$EMNO, PN.T$NAMA, Project, [week], Jaar");
            query.AppendLine("Order by PN.T$EMNO, UR.Jaar, UR.[week], PN.T$NAMA");

            List<Worker> returnValue = new List<Worker>();
            using (IADOConnection conn = ADOFactory.Create() as IADOConnection)
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = query.ToString();
                    SqlDataReader reader = (SqlDataReader)cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        Worker worker = new Worker();

                        worker.Number = Convert.ToInt32(reader["PersoneelNummer"]);
                        worker.Name = reader["Name"].ToString();
                        worker.ProjectNumber = reader["Project"].ToString();
                        worker.Year = Convert.ToInt32(reader["Jaar"]);
                        worker.Week = Convert.ToInt32(reader["Week"]);
                        worker.Hours = Convert.ToDouble(reader["Hours"]);

                        returnValue.Add(worker);

                    }

                    reader.Close();
                    reader = null;

                }
            }
            return returnValue;


        }


    }
}
