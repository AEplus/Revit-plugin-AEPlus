using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using xlApp = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using OfficeOpenXml;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CSVPerTabToevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Result r = Result.Failed;
            xlApp.Application ExcelApp = null;
            xlApp.Workbook workbook = null;
            xlApp.Sheets sheets = null;
            xlApp.Worksheet newSheet = null;

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
            ViewScheduleExportOptions opt = new ViewScheduleExportOptions
            {
                TextQualifier = ExportTextQualifier.None,
                FieldDelimiter = ","
            };

            //Instantiate the Application object.
            xlApp.Application excelApp = new xlApp.Application();

            //Open the excel file.
            workbook = excelApp.Workbooks.Open(@"c:\\temp\\test.xlsx", 0, false, 5, "", "",
                           false, XlPlatform.xlWindows, "",
                           true, false, 0, true, false, false);

            //Declare a Worksheet object.
            sheets = workbook.Sheets as Sheets;

            int i = 1;
            foreach (ViewSchedule vs in col)
            {
                vs.Export(@"c:\\temp\\test", Environment.UserName + vs.Name + ".csv", opt);

                newSheet = (Worksheet)sheets.Add(sheets[i], Type.Missing, Type.Missing, Type.Missing);




                // workbook.ActiveSheet(@"c:\\temp\\test\" + vs.Name + ".csv");
                //   newSheet.Name = vs.Name;
                //   workbook.Save();



            }

            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                ExcelApp.Quit();
                Marshal.ReleaseComObject(newSheet);
                Marshal.ReleaseComObject(sheets);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                return r;
           
            return r;
        }
    } 
}
