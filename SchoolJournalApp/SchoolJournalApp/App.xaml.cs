using OfficeOpenXml;
using System.Windows;

namespace SchoolJournalApp
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Явно указываем пространство имён OfficeOpenXml
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            base.OnStartup(e);
        }
    }
}