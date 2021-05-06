using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

// Baby version of generic toevoegen 
// This programs is the first item in the ribbon and adds every schedule containing the name "AE"
// You should watch if the schedules names are not equal under 31chars. This gives an error.
// It writes to c:\temp which it creates, writing and reading is possible on the server
// For securiy purposes this is only done locally so nobody can tinker with this on the server

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class TotaalToevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string xlSheetName;
            Directory.CreateDirectory(@"c:\\temp\\totaal");
            var r = Result.Succeeded;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var MapPath = @"c:\\temp\\totaal\";

            // MisValue for excel files, unknown values
            object misValue = Missing.Value;

            // Get UIDocument
            var uidoc = commandData.Application.ActiveUIDocument;
            // Get Document
            var doc = uidoc.Document;

            var col = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule));

            // Export Options on how the .txt file will look. Text "" is gone
            // FieldDelimiter is TAB replaced with ,
            var opt = new ViewScheduleExportOptions
            {
                TextQualifier = ExportTextQualifier.DoubleQuote,
                FieldDelimiter = ","
            };
            using (var excelEngine = new ExcelPackage())
            {
                // Foreach mag hier pas staan want maakt anders per viewschedule een overwrite van het excelpackage
                // waardoor je geen totaaloverzicht kan krijgen
                foreach (ViewSchedule vs in col)
                {
                    var fileName = Environment.UserName + vs.Name;

                    // Searches for schedules containing AE
                    if (vs.Name.Contains("AE"))
                    {
                        // Checks if name length is more than 31 because Excel sheets can not contain more characters.
                        if (vs.Name.Length > 30)
                            xlSheetName = vs.Name.Substring(0, 30);
                        else
                            xlSheetName = vs.Name;
                        //create a WorkSheet
                        var ws = excelEngine.Workbook.Worksheets.Add(xlSheetName);

                        // Export c:\\temp --> Will be save as
                        vs.Export(MapPath, fileName + ".csv", opt);
                        var file = new FileInfo(MapPath + fileName + ".csv");

                        // Adds Worksheet as first in the row
                        ws.Workbook.Worksheets.MoveToStart(xlSheetName);

                        // Formating for writing to xlsx
                        var format = new ExcelTextFormat
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
                        var filePath = MapPath + "ExcelSchedulesTotaal.xlsx";

                        // Write the file to the disk
                        var fi = new FileInfo(filePath);
                        excelEngine.SaveAs(fi);

                        // Deletes every .csv file so only the .xslx stays in the folder.
                        File.Delete(MapPath + fileName + ".csv");
                    }
                }
            }
            return r;
        }
    }
}