using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using WinForms = System.Windows.Forms;
using X = Microsoft.Office.Interop.Excel;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class ExportParameter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiApp = commandData.Application;
            var app = uiApp.Application;
            var doc = uiApp.ActiveUIDocument.Document;

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

            using (var excelEngine = new ExcelPackage())
            {
                var wsParameters = excelEngine.Workbook.Worksheets.Add("Parameters");

                wsParameters.Cells[1, 1].Value = "Family";
                wsParameters.Cells[1, 2].Value = "FamilySymbol";
                wsParameters.Cells[1, 3].Value = "GeoIT Value";
                wsParameters.Cells[1, 4].Value = "Unique ID";

                var fs
                    = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();

                // var fs = families.WhereElementIsElementType().ToElements();
                var rij = 2;

                foreach (var e in fs)
                {
                    wsParameters.Cells[rij, 1].Value = e.FamilyName;
                    wsParameters.Cells[rij, 2].Value = e.Name;
                    if (e.get_Parameter(guid) != null)
                        wsParameters.Cells[rij, 3].Value = e.get_Parameter(guid).AsValueString();
                    else
                        wsParameters.Cells[rij, 3].Value = "";
                    wsParameters.Cells[rij, 4].Value = e.UniqueId;

                    rij++;
                }

                excelEngine.SaveAs(new FileInfo(@"c:\temp\GeoITParameter.xlsx"));
            }

            // workbook.SaveAs(workbook, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            //    X.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            //    Missing.Value);

            return Result.Succeeded;
        }
    }
}