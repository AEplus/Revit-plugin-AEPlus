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
                // Gets GUID of AE GeoIT Link parameter. 
                // This is done for because name could cause errors. 
                var file = app.OpenSharedParameterFile();
                var group = file.Groups.get_Item("00. General");
                var definition = group.Definitions.get_Item("AE GeoIT Link");
                var externalDefinition = definition as ExternalDefinition;
                guid = externalDefinition.GUID;
            }
            catch (Exception)
            {
            }

            using (var excelEngine = new ExcelPackage())
            {
                var wsParameters = excelEngine.Workbook.Worksheets.Add("Parameters");

                // Add headers for excel file 
                wsParameters.Cells[1, 1].Value = "Family";
                wsParameters.Cells[1, 2].Value = "FamilySymbol";
                wsParameters.Cells[1, 3].Value = "GeoIT Value";

                // Filtering for every family type that is present in the current project. 
                var familyTypeCollector1 =
                    new FilteredElementCollector(doc);
                familyTypeCollector1.OfClass(typeof(ElementType));
                var famTypes = familyTypeCollector1.ToElements();

                // Excel starts counting at 1 so the first row with data filled in is 2
                var row = 2;


                // start transaction 
                using (var t = new Transaction(doc))
                {
                    t.Start("Symbol activate");

                    // Loops through all the family types that are present
                    foreach (var e in famTypes)
                    {
                        // Changes element to elementType for access to parameters.
                        var type = e as ElementType;

                        // Got builtinparameters value through RevitLookUp .AsString is necessary to get value instead of DB.Parameter
                        wsParameters.Cells[row, 1].Value =
                            type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString();
                        wsParameters.Cells[row, 2].Value =
                            type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString();
                        var number = 0;

                        // Parameter can be empty because some system families and types do not allow being changed
                        if (e.get_Parameter(guid) != null)
                        {
                            // Changes excel from standard to number, this is done for importing.
                            int.TryParse(e.get_Parameter(guid).AsValueString(), out number);
                            wsParameters.Cells[row, 3].Value = number;
                            wsParameters.Cells["C1:C" + wsParameters.Dimension.End.Row].Style.Numberformat.Format = "0";
                        }
                        else
                        {
                            // if null fill in 0
                            wsParameters.Cells[row, 3].Value = 0;
                        }

                        row++;
                    }

                    t.Commit();
                }

                var fi = "";

                // SaveFileDialog to save excel
                var dlg = new WinForms.SaveFileDialog();
                dlg.InitialDirectory = @"C:\";
                dlg.Title = "Select directory to save the excel file";
                dlg.Filter = "Excel spreadsheet files (*.xlsx)|*.xlsx|All files (*)|*";

                // save file ok 
                if (dlg.ShowDialog() == WinForms.DialogResult.OK) fi = dlg.FileName;

                try
                {
                    excelEngine.SaveAs(new FileInfo(fi));
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e);
                    return Result.Cancelled;
                }
            }

            return Result.Succeeded;
        }
    }
}