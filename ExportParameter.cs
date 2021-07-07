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

                var row = 2;
                var cellValue = 0;

                using (var t = new Transaction(doc))
                {
                    t.Start("Symbol activate");

                    foreach (var e in fs)
                    {
                        //if (!e.IsActive)x 
                        //{
                        //    e.Activate();
                        //    doc.Regenerate();
                        //}

                        wsParameters.Cells[row, 1].Value = e.FamilyName;
                        wsParameters.Cells[row, 2].Value = e.Name;
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