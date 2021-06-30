using System;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

            var families
                = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol));

            //var excel = new X.Application();

            //excel.Visible = true;
            //var workbook = excel.Workbooks.Add(Missing.Value);
            //X.Worksheet worksheet;

            //worksheet = excel.ActiveSheet as X.Worksheet;
            //worksheet.Name = "GeoIT Link";
            //worksheet.Cells[1, 1] = "FamilySymbol";
            //worksheet.Cells[1, 2] = "Family";
            //worksheet.Cells[1, 3] = "GeoIT Link";

            var row = 2;

            var test = fs. wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww

            foreach (FamilySymbol fs in families)
            {
                var geoit = fs.Symbol.LookupParameter("AE GeoIT Link");
                var typeName = fs.Name;
                var familyName = fs.FamilyName;


                Debug.WriteLine(geoit.ToString() + typeName + familyName + Environment.NewLine);
                
                //worksheet.Cells[row, 1] = typeName;
                //worksheet.Cells[row, 2] = familyName;

                //if (link == null)
                //    worksheet.Cells[row, 3] = 0;
                //else
                //    worksheet.Cells[row, 3] = link;


                row++;
            }

            // workbook.SaveAs(workbook, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            //    X.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
            //    Missing.Value);

            return Result.Succeeded;
        }
    }
}