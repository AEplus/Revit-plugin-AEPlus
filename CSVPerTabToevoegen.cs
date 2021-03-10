using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using OfficeOpenXml;
using System.IO;
using System.Globalization;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CSVPerTabToevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.IO.Directory.CreateDirectory(@"c:\\temp\");
            Result r = Result.Succeeded;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // MisValue for excel files, unknown values
            object misValue = System.Reflection.Missing.Value;

            // Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            // Get Document
            Document doc = uidoc.Document;

            FilteredElementCollector col = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSchedule));
            
           // Export Options on how the .txt file will look. Text "" is gone
           // FieldDelimiter is TAB replaced with ,            
            ViewScheduleExportOptions opt = new ViewScheduleExportOptions()
            {
                TextQualifier = ExportTextQualifier.DoubleQuote,
                FieldDelimiter = ","
            };

            using (ExcelPackage excelEngine = new ExcelPackage())
            {
                foreach (ViewSchedule vs in col)
                {
                    //create a WorkSheet
                    ExcelWorksheet ws = excelEngine.Workbook.Worksheets.Add(vs.Name);

                    // Searches for schedules containing AE 
                    if (vs.Name.Contains("AE"))
                    {
                        // Export c:\\temp --> Will be save as
                        vs.Export(@"c:\\temp\", Environment.UserName + vs.Name + ".csv", opt);
                        FileInfo file = new FileInfo(@"c:\\temp\" + Environment.UserName + vs.Name + ".csv");

                        // Adds Worksheet as first in the row 
                        ws.Workbook.Worksheets.MoveToStart(vs.Name);

                        // Formating for writing to xlsx
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
                        // the path of the file
                        string filePath = "C:\\temp\\ExcelSchedulesTotaal.xlsx";

                        // Write the file to the disk
                        FileInfo fi = new FileInfo(filePath);
                        excelEngine.SaveAs(fi);

                        
                        // Create SaveFileDialog here
                        // Future project
                    }
                }
            }
            return r;
        }
    }
}
