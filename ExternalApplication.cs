using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;

// Dit is het algemeen exporteren van alle schedules
// V
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
            var strImageFolder = Path.GetDirectoryName(path) + @"\Resources\";

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

            var buttonElekAantal = new PushButtonData("elek aantal",
                "Export met aantallen",
                path,
                "MyRevitCommands.ElekAantal");


            var buttonElekLengte = new PushButtonData("Elek lengte",
                "Export met lengte",
                path,
                "MyRevitCommands.ElekLengte");


            var buttonHVACAantal = new PushButtonData("HVAC Aantal",
                "Export met aantallen",
                path,
                "MyRevitCommands.HVACAantal");

            var buttonHVACLengte = new PushButtonData("Export HVAC lengte",
                "Export met lengte",
                path,
                "MyRevitCommands.HVACLengte");

            var buttonSanAantal = new PushButtonData("Export Sanitair Aantallen",
                "Export met aantallen",
                path,
                "MyRevitCommands.SanAantal");

            var buttonSanLengte = new PushButtonData("Export Sanitair lengte",
                "Export met lengte",
                path,
                "MyRevitCommands.SanLengte");


            var buttonUpdater = new PushButtonData("Updater",
                "Updater",
                path,
                "MyRevitCommands.Updater");

            var buttonExport = new PushButtonData("Export parameters",
                "Export parameters",
                path,
                "MyRevitCommands.ExportParameter");

            var buttonImport = new PushButtonData("Import parameters",
                "Import parameters",
                path,
                "MyRevitCommands.ImportParameter");

            //var buttonTest = new PushButtonData("Test",
            //    "test",
            //    path,
            //    "MyRevitCommands.Test");



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

            //panelTest.AddItem(buttonTest);


            return Result.Succeeded;
        }
    }
}