using System;
using System.IO;
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

                var familyTypeCollector1 =
                    new FilteredElementCollector(doc);
                familyTypeCollector1.OfClass(typeof(ElementType));
                var famTypes = familyTypeCollector1.ToElements();

                var row = 2;
                var cellValue = 0;

                using (var t = new Transaction(doc))
                {
                    t.Start("Symbol activate");

                    foreach (var e in famTypes)
                    {
                        var type = e as ElementType;
                        //if (!e.IsActive)x 
                        //{
                        //    e.Activate();
                        //    doc.Regenerate();
                        //}

                        // Got builtinparameters value through RevitLookUp .AsString is necessary to get value instead of DB.Parameter
                        wsParameters.Cells[row, 1].Value =
                            type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString();
                        wsParameters.Cells[row, 2].Value =
                            type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString();
                        var number = 0;

                        if (e.get_Parameter(guid) != null)
                        {
                            int.TryParse(e.get_Parameter(guid).AsValueString(), out number);
                            wsParameters.Cells[row, 3].Value = number;
                            wsParameters.Cells["C1:C" + wsParameters.Dimension.End.Row].Style.Numberformat.Format = "0";
                        }
                        else
                        {
                            wsParameters.Cells[row, 3].Value = 0;
                        }

                        row++;
                    }

                    t.Commit();
                }

                var fi = "";

                var dlg = new WinForms.SaveFileDialog();
                dlg.InitialDirectory = @"C:\";
                dlg.Title = "Select directory to save the excel file";
                dlg.Filter = "Excel spreadsheet files (*.xlsx)|*.xlsx|All files (*)|*";
                if (dlg.ShowDialog() == WinForms.DialogResult.OK) fi = dlg.FileName;

                excelEngine.SaveAs(new FileInfo(fi));
            }

            return Result.Succeeded;
        }
    }
}