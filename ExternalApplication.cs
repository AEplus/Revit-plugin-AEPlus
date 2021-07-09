using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

// Alles op startup voor opstarten van de plugin
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
            // Gets the assembly directory 
            var path = Assembly.GetExecutingAssembly().Location;
           
            var AEPlusLogoico =
                new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/AEPLusLogo.ico"));

            var numbersico =
                new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/numbers.ico"));

            var lengthico =
                new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/length.ico"));

            var updaterico =
                new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/updater.ico"));
            var importico =
                new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/import.ico"));
            var exportico =
                new BitmapImage(new Uri("pack://application:,,,/MyRevitCommands;component/Resources/export.ico"));

            // Create a custom ribbon tab
            var tabName = "AE Plus";
            application.CreateRibbonTab(tabName);

            // Create a ribbon panel
            var panelTot = application.CreateRibbonPanel(tabName, "Totaal");
            var panelHvac = application.CreateRibbonPanel(tabName, "Export HVAC");
            var panelSan = application.CreateRibbonPanel(tabName, "Export Sanitair");
            var panelElek = application.CreateRibbonPanel(tabName, "Export elektriciteit");
            var panelUpdate = application.CreateRibbonPanel(tabName, "Updater");

            // Create button
            var buttonTot = new PushButtonData("Export schedules",
                "Export schedules",
                path,
                "MyRevitCommands.TotaalToevoegen");
            buttonTot.LargeImage = AEPlusLogoico;

            var buttonElekAantal = new PushButtonData("elek aantal",
                "Export met aantallen",
                path,
                "MyRevitCommands.ElekAantal");
            buttonElekAantal.LargeImage = numbersico;

            var buttonElekLengte = new PushButtonData("Elek lengte",
                "Export met lengte",
                path,
                "MyRevitCommands.ElekLengte");
            buttonElekLengte.LargeImage = lengthico;

            var buttonHVACAantal = new PushButtonData("HVAC Aantal",
                "Export met aantallen",
                path,
                "MyRevitCommands.HVACAantal");
            buttonHVACAantal.LargeImage = numbersico;

            var buttonHVACLengte = new PushButtonData("Export HVAC lengte",
                "Export met lengte",
                path,
                "MyRevitCommands.HVACLengte");
            buttonHVACLengte.LargeImage = lengthico;

            var buttonSanAantal = new PushButtonData("Export Sanitair Aantallen",
                "Export met aantallen",
                path,
                "MyRevitCommands.SanAantal");
            buttonSanAantal.LargeImage = numbersico;

            var buttonSanLengte = new PushButtonData("Export Sanitair lengte",
                "Export met lengte",
                path,
                "MyRevitCommands.SanLengte");
            buttonSanLengte.LargeImage = lengthico;

            var buttonUpdater = new PushButtonData("Updater",
                "Updater",
                path,
                "MyRevitCommands.Updater");
            buttonUpdater.LargeImage = updaterico;

            var buttonExport = new PushButtonData("Export parameters",
                "Export parameters",
                path,
                "MyRevitCommands.ExportParameter");
            buttonExport.LargeImage = exportico;

            var buttonImport = new PushButtonData("Import parameters",
                "Import parameters",
                path,
                "MyRevitCommands.ImportParameter");
            buttonImport.LargeImage = importico;

            // Adding items to the panel
            panelTot.AddItem(buttonTot);

            panelElek.AddItem(buttonElekAantal);
            panelElek.AddItem(buttonElekLengte);

            panelHvac.AddItem(buttonHVACAantal);
            panelHvac.AddItem(buttonHVACLengte);

            panelSan.AddItem(buttonSanAantal);
            panelSan.AddItem(buttonSanLengte);

            panelUpdate.AddItem(buttonUpdater);

            panelUpdate.AddItem(buttonExport);
            panelUpdate.AddItem(buttonImport);

            return Result.Succeeded;
        }
    }
}