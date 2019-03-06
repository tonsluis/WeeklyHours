
namespace WeeklyHoursXlReportBuilder
{
    #region Directives
    // Standard .NET Directives
    using System;
    using System.Configuration;
    #endregion










    /// <summary>
    /// Neger object
    /// </summary>
    public class Neger : ConfigurationElement 
    {
        [ConfigurationProperty("Id",IsKey =true,  IsRequired = true)]
        public int Id
        {
            get
            {
                return (int)this["Id"];
            }
            set
            {
                value = (int)this["Id"];
            }
        }

        [ConfigurationProperty("Name", DefaultValue = "Tom", IsRequired = true)]
        public string Name
        {
            get
            {
                return (string)this["Name"];
            }
            set
            {
                value = (string)this["Name"];
            }
        }

    }
}
