using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using OfficeOpenXml;
using System.IO;

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

            foreach (ViewSchedule vs in col)
            {
                if (vs.Name.Contains("AE"))
                {
                    vs.Export(@"c:\\temp\", Environment.UserName + vs.Name + ".csv", opt);


                    FileInfo file = new FileInfo(@"c:\\temp\" + Environment.UserName + vs.Name + ".csv");

                    //create a new Excel package
                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        //create a WorkSheet
                        ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add(vs.Name);

                        ExcelTextFormat format = new ExcelTextFormat();

                        //load the CSV data into cell A1
                        worksheet.Cells["A1"].LoadFromText(file, format);

                        //the path of the file
                        string filePath = "C:\\temp\\ExcelDemo.xlsx";

                        //Write the file to the disk
                        FileInfo fi = new FileInfo(filePath);
                        excel.SaveAs(fi);
                    }
                }
            }
            //newSheet = (Worksheet)sheets.Add(sheets[i], Type.Missing, Type.Missing, Type.Missing);

            // workbook.ActiveSheet(@"c:\\temp\\test\" + vs.Name + ".csv");
            //   newSheet.Name = vs.Name;
            //   workbook.Save();

            // Marshal.ReleaseComObject(newSheet);
            // Marshal.ReleaseComObject(sheets);
            // Marshal.ReleaseComObject(workbook);

            return r;
        }
    }
}