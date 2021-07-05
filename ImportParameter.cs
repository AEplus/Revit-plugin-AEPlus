using System;
using System.Diagnostics;
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
            if (WinForms.DialogResult.OK != dlg.ShowDialog()) return Result.Failed;


            var filename = new FileInfo(dlg.FileName);
            var filepath = Path.GetDirectoryName(dlg.FileName);


            using (var excelEngine = new ExcelPackage(filename))
            {
                var workbook = excelEngine.Workbook;
                var worksheet = excelEngine.Workbook.Worksheets[1];

                using (var t = new Transaction(doc))
                {
                    if (t.Start("Shared parameter values invullen") == TransactionStatus.Started)
                    {
                        var families
                            = new FilteredElementCollector(doc)
                                .OfClass(typeof(FamilySymbol));

                        var fs = families.WhereElementIsElementType().ToElements();

                        foreach (FamilySymbol type in fs)
                        {
                            var rij = 2;

                            try
                            {
                                var famname = type.FamilyName;
                                var famsymbol = type.Name;
                                var excelfamily = worksheet.Cells[rij, 1].Value;
                                var excelfs = worksheet.Cells[rij, 2].Value;
                                var gvalue = (double) worksheet.Cells[rij, 3].Value;


                                if (Equals(famname, excelfamily))
                                    if (Equals(famsymbol, excelfs))
                                    {
                                        var fspara = type.get_Parameter(guid);
                                        fspara.Set(gvalue);
                                        Debug.WriteLine(famname + famsymbol);
                                        Debug.WriteLine(gvalue);
                                    }


                                Console.WriteLine("Volgende familysymbol");
                            }

                            catch (Exception exception)
                            {
                                Debug.WriteLine(exception);
                            }

                            rij++;
                        }

                        t.Commit();
                    }
                }

                return Result.Succeeded;
            }
        }
    }
}