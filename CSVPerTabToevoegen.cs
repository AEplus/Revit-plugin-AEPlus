using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Globalization;

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
                TextQualifier = ExportTextQualifier.DoubleQuote,
                FieldDelimiter = ","
            };

            int i = 1;
            foreach (ViewSchedule vs in col)
            {
                using (ExcelPackage excelEngine = new ExcelPackage())
                {
                if (vs.Name.Contains("AE"))
                {
                    vs.Export(@"c:\\temp\", Environment.UserName + vs.Name + ".csv", opt);

                        FileInfo file = new FileInfo(@"c:\\temp\" + Environment.UserName + vs.Name + ".csv");

                        //create a WorkSheet
                        ExcelWorksheet ws = excelEngine.Workbook.Worksheets.Add(vs.Name);
                        int totalsheets = excelEngine.Workbook.Worksheets.Count;

                        var format = new ExcelTextFormat()
                        {   
                          Culture = CultureInfo.InvariantCulture,
                        //    // Escape character for values containing the Delimiter
                        //    // ex: "A,Name",1 --> two cells, not three
                        TextQualifier = '"'
                        //    // Other properties
                        //    // EOL, DataTypes, Encoding, SkipLinesBeginning/End
                        };
                        ws.Cells["A1"].LoadFromText(file, format);
                        
                       // excelEngine.Workbook.Worksheets.MoveBefore(i);

                        //the path of the file
                        string filePath = "C:\\temp\\ExcelDemo.xlsx";

                        //Write the file to the disk
                        FileInfo fi = new FileInfo(filePath);
                        excelEngine.SaveAs(fi);
                    }
                }
            }
            return r;
        }
    }
}
