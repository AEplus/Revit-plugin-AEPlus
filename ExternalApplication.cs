using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace MyRevitCommands
{
    internal class ExternalApplication : IExternalApplication
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

            // Create a ribbon panel
            RibbonPanel panelTot = application.CreateRibbonPanel(tabName, "Totaal");
            RibbonPanel panelHvac = application.CreateRibbonPanel(tabName, "Export HVAC");
            RibbonPanel panelSan = application.CreateRibbonPanel(tabName, "Export Sanitair");
            RibbonPanel panelElek = application.CreateRibbonPanel(tabName, "Export elektriciteit");

            // Create button
            PushButtonData buttonTot = new PushButtonData("Export schedules",
                                                          "Export schedules",
                                                          path,
                                                          "MyRevitCommands.TotaalToevoegen");

            PushButtonData buttonElekAantal = new PushButtonData("elek aantal",
                                                        "Export met aantallen",
                                                        path,
                                                        "MyRevitCommands.ElekAantal");


            PushButtonData buttonElekLengte = new PushButtonData("Elek lengte",
                                                        "Export met lengte",
                                                        path,
                                                        "MyRevitCommands.ElekLengte");


            PushButtonData buttonHVACAantal = new PushButtonData("HVAC Aantal",
                                                                 "Export met aantallen",
                                                                 path,
                                                                 "MyRevitCommands.HVACAantal");

            PushButtonData buttonHVACLengte = new PushButtonData("Export HVAC lengte",
                                                                 "Export met lengte",
                                                                 path,
                                                                 "MyRevitCommands.HVACLengte");

            PushButtonData buttonSanAantal = new PushButtonData("Export Sanitair Aantallen",
                                                                "Export met aantallen",
                                                                path,
                                                                "MyRevitCommands.SanAantal");

            PushButtonData buttonSanLengte = new PushButtonData("Export Sanitair lengte",
                                                                "Export met lengte",
                                                                path,
                                                                "MyRevitCommands.SanLengte");

            BitmapImage button1image = new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/favicon.ico"));
            buttonTot.LargeImage = button1image;

            BitmapImage button2image = new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/numbers.ico"));
            buttonElekAantal.LargeImage = button2image;

            BitmapImage button3image = new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/pipe.ico"));
            buttonSanLengte.LargeImage = button3image;

            panelTot.AddItem(buttonTot);

            panelElek.AddItem(buttonElekAantal);
            panelElek.AddItem(buttonElekLengte);

            panelHvac.AddItem(buttonHVACAantal);
            panelHvac.AddItem(buttonHVACLengte);

            panelSan.AddItem(buttonSanAantal);
            panelSan.AddItem(buttonSanLengte);

            return Result.Succeeded;
        }
    }
}