using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace MyRevitCommands
{
    class ExternalApplication : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
            string path = Assembly.GetExecutingAssembly().Location;
            string strImageFolder = System.IO.Path.GetDirectoryName(path) + @"\Resources\";

            // Create a custom ribbon tab
            String tabName = "AE Plus";
            application.CreateRibbonTab(tabName);

            // Create button           
            PushButtonData button1 = new PushButtonData("Export schedules" , "Export schedules", path, "MyRevitCommands.CSVPerTabToevoegen");

            // Create a ribbon panel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Export Schedules");

            panel.AddItem(button1);

            return Result.Succeeded;
        }
    }
}
