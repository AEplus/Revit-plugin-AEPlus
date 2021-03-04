using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using OfficeOpenXml;
using System.IO;
using System.Linq;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CSVPerTabToevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Result r = Result.Succeeded;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // MisValue for excel files, unknown values
            object misValue = System.Reflection.Missing.Value;

            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get Document
            Document doc = uidoc.Document;

            FilteredElementCollector col = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSchedule));

            /*
           Export Options on how the .txt file will look. Text "" is gone
           FieldDelimiter is TAB replaced with ,  
            */
            ViewScheduleExportOptions opt = new ViewScheduleExportOptions()
            {
                TextQualifier = ExportTextQualifier.DoubleQuote
            };

            int i = 0;
            foreach (ViewSchedule vs in col)
            {
                if (vs.Name.Contains("AE"))
                {
                    vs.Export(@"c:\\temp\", Environment.UserName + vs.Name + ".csv", opt);

                    FileInfo file = new FileInfo(@"c:\\temp\" + Environment.UserName + vs.Name + ".csv");

                    using (ExcelPackage excelEngine = new ExcelPackage())
                    {
                        var xlWorkbook = excelEngine.Workbook.Worksheets.Add(vs.Name);
                        xlWorkbook.Workbook.Worksheets.Count();
                        xlWorkbook.Workbook.Worksheets.MoveToEnd(i);
                        i++;
                    }


                }
            }
            return r;
        }
    }
}
