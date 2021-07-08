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

            var file = app.OpenSharedParameterFile();
            var group = file.Groups.get_Item("00. General");
            var definition = group.Definitions.get_Item("AE GeoIT Link");
            var externalDefinition = definition as ExternalDefinition;
            guid = externalDefinition.GUID;

            var dlg = new WinForms.OpenFileDialog();
            dlg.Title = "Select source Excel file from which to update Revit shared parameters";
            dlg.Filter = "Excel spreadsheet files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*)|*";
            if (WinForms.DialogResult.OK != dlg.ShowDialog()) return Result.Failed;
            var filename = new FileInfo(dlg.FileName);

            using (var excelEngine = new ExcelPackage(filename))
            {
                worksheet = excelEngine.Workbook.Worksheets[1];
                var iRowCnt = worksheet.Dimension.End.Row;

                using (var t = new Transaction(doc))
                {
                    t.Start("started");

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
            var fstype = type;
            string famname = null;
            string famsymbol = null;

            // if (!type.IsActive) type.Activate();

            if (type != null)
            {
                famname = type.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString();
                famsymbol = type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString();
            }

            var run = true;
            while (run && row < iRowCnt)
            {
                var excelfamily = (string) worksheet.Cells[row, 1].Value;
                var excelfs = (string) worksheet.Cells[row, 2].Value;
                var gvalue = (double) worksheet.Cells[row, 3].Value;

                if (famname.Equals(excelfamily))
                    if (famsymbol.Equals(excelfs))
                    {
                        var fspara = type.get_Parameter(guid);

                        if (fspara != null && fspara.IsReadOnly != true)
                        {
                            fspara.Set(gvalue);
                            //Debug.WriteLine("OK");
                        }

                        run = false;
                    }

                row++;
            }


/*
            var excelfamily = (string) worksheet.Cells[row, 1].Value;
            var excelfs = (string) worksheet.Cells[row, 2].Value;
            var gvalue = (double) worksheet.Cells[row, 3].Value;

            var skip = false;
            if (famname.Equals(excelfamily))
                if (famsymbol.Equals(excelfs))
                {
                    var fspara = type.get_Parameter(guid);
                    if (fspara == null)
                        skip = true;
                    else if (fspara.IsReadOnly != true)
                        fspara.Set(gvalue);
                }

            if (!skip && row + 1 < iRowCnt && (famname != excelfamily || famsymbol != excelfs))
                checkFamily(row + 1, iRowCnt, type);
*/
        }
    }
}