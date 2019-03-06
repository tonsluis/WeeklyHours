
namespace WeeklyHoursXlReportBuilder
{
    #region Directives
    // Standard .NET Directives
    using System;
    using System.Configuration;
    #endregion

    /// <summary>
    /// ReportSubjects object
    /// </summary>
    public class ReportSubjects : ConfigurationSection 
    {

        [ConfigurationProperty("",IsRequired =true,IsDefaultCollection = true)]
        public Negers Instances
        {
            get { return (Negers)this[""]; }
            set { this[""] = value; }
        }
    }


    public class Negers : ConfigurationElementCollection
    {

        protected override ConfigurationElement CreateNewElement()
        {
            return new Neger();
        }


        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((Neger)element).Id;
        }

    }




}
