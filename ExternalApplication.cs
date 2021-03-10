﻿using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Media.Imaging;

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
            PushButtonData button1 = new PushButtonData("Export schedules", "Export schedules", path, "MyRevitCommands.CSVPerTabToevoegen");
            PushButtonData button2 = new PushButtonData("Export E60", "Export E60", path, "E_60Toevoegen.CSV_E60Toevoegen");

            BitmapImage button1image = new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/favicon.ico"));
            button1.LargeImage = button1image;

            // Create a ribbon panel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Export Schedules");

            panel.AddItem(button1);
            panel.AddItem(button2);
            return Result.Succeeded;
        }
    }
}