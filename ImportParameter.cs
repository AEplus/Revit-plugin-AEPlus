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
        private Guid guid;
        private ExcelWorksheet worksheet;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var app = uiapp.Application;
            var doc = uiapp.ActiveUIDocument.Document;
            var row = 2;

            // Gets shared parameter through guid 
            var file = app.OpenSharedParameterFile();
            var group = file.Groups.get_Item("00. General");
            var definition = group.Definitions.get_Item("AE GeoIT Link");
            var externalDefinition = definition as ExternalDefinition;
            guid = externalDefinition.GUID;

            // OpenFileDialog to choose excel file
            var dlg = new WinForms.OpenFileDialog();
            dlg.Title = "Select source Excel file from which to update Revit shared parameters";
            dlg.Filter = "Excel spreadsheet files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*)|*";
            if (WinForms.DialogResult.OK != dlg.ShowDialog()) return Result.Failed;
            var filename = new FileInfo(dlg.FileName);

            using (var excelEngine = new ExcelPackage(filename))
            {
                // Only one worksheet is present. 
                worksheet = excelEngine.Workbook.Worksheets[1];

                // End row count, number to when stop looping
                var iRowCnt = worksheet.Dimension.End.Row;

                using (var t = new Transaction(doc))
                {
                    t.Start("started");

                    // Same filter as export, loops through every element(type)
                    var familyTypeCollector1 = new FilteredElementCollector(doc);
                    familyTypeCollector1.OfClass(typeof(ElementType));
                    var famTypes = familyTypeCollector1.ToElements();

                    foreach (var type in famTypes)

                        checkFamily(iRowCnt, (ElementType) type);

                    t.Commit();
                    return Result.Succeeded;
                }
            }
        }

        private void checkFamily(int iRowCnt, ElementType type)
        {
            // Debug.WriteLine("BEGIN");
            var row = 2;
            string famname = null;
            string famsymbol = null;

            if (type != null)
            {
                // Uses builtinparameters to not encounter errors or other parameters values with the same names
                famname = type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString();
                famsymbol = type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString();
            }

            var run = true;
            while (run && row < iRowCnt)
            {
                // Reads excel values for each row
                var excelfamily = (string) worksheet.Cells[row, 1].Value;
                var excelfs = (string) worksheet.Cells[row, 2].Value;
                var gvalue = (double) worksheet.Cells[row, 3].Value;

                // Compares actual famname and fs name to excel values 
                if (famname.Equals(excelfamily))
                    if (famsymbol.Equals(excelfs))
                    {
                        var fspara = type.get_Parameter(guid);

                        if (fspara != null && fspara.IsReadOnly != true)
                            // Writes value
                            fspara.Set(gvalue);

                        run = false;
                    }

                row++;
            }
        }
    }
}