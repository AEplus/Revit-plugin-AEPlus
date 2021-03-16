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
        string xlSheetName;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.IO.Directory.CreateDirectory(@"c:\\temp\\totaal");
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
                // Foreach mag hier pas staan want maakt anders per viewschedule een overwrite van het excelpackage
                // waardoor je geen totaaloverzicht kan krijgen
                foreach (ViewSchedule vs in col)
                {  
                    // Searches for schedules containing AE 
                    if (vs.Name.Contains("AE"))
                    {
                        // Checks if name length is more than 31 because Excel sheets can not contain more characters.
                        if (vs.Name.Length > 30)
                        {
                            xlSheetName = vs.Name.Substring(0, 30);
                        }
                        else
                        {
                            xlSheetName = vs.Name;
                        }
                            
                        //create a WorkSheet
                        ExcelWorksheet ws = excelEngine.Workbook.Worksheets.Add(xlSheetName);

                        // Export c:\\temp --> Will be save as
                        vs.Export(@"c:\\temp\\totaal\", Environment.UserName + vs.Name + ".csv", opt);
                        FileInfo file = new FileInfo(@"c:\\temp\\totaal\\" + Environment.UserName + vs.Name + ".csv");

                        // Adds Worksheet as first in the row 
                        ws.Workbook.Worksheets.MoveToStart(xlSheetName);

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
                        string filePath = "C:\\temp\\totaal\\ExcelSchedulesTotaal.xlsx";

                        // Write the file to the disk
                        FileInfo fi = new FileInfo(filePath);
                        excelEngine.SaveAs(fi);
                    }
                }
            }
            return r;
        }
    }
}

