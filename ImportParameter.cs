using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using WinForms = System.Windows.Forms;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class ImportParameter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var app = uiapp.Application;
            var doc = uiapp.ActiveUIDocument.Document;

            var guid = Guid.Empty;
            try
            {
                var file = app.OpenSharedParameterFile();
                var group = file.Groups.get_Item("00. General");
                var definition = group.Definitions.get_Item("AE GeoIT Link");
                var externalDefinition = definition as ExternalDefinition;
                guid = externalDefinition.GUID;
                Console.WriteLine(guid);
            }
            catch (Exception)
            {
            }

            var dlg = new WinForms.OpenFileDialog();
            dlg.Title = "Select source Excel file from which to update Revit shared parameters";
            dlg.Filter = "Excel spreadsheet files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*)|*";
            if (WinForms.DialogResult.OK != dlg.ShowDialog()) return Result.Cancelled;

            var newFile = new FileInfo(dlg.ToString());

            using (var excelEngine = new ExcelPackage(newFile))
            {
                var workbook = excelEngine.Workbook;
                var worksheet = excelEngine.Workbook.Worksheets[1];


                using (var t = new Transaction(doc))
                {
                    t.Start("Shared parameter values invullen");
                    var rij = 2;

                    while (true)
                    {
                        //foreach (symbolId as ElementId In family.GetFamilySymbolIds())
                        {
                            //var symbol as FamilySymbol = doc.GetElement(symbolId);
                        }

                        // if symbol.name.Equals(desiredFamilySymbolName, StringComparison.InvariantCultureIgnoreCase)


                        break;
                    }
                }
            }

            return Result.Succeeded;
        }
    }
}