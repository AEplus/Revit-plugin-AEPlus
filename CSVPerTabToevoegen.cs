using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using xlApp = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using OfficeOpenXml;
using System.Text;
using System.Globalization;
using System.Threading;
using Microsoft.IO;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CSVPerTabToevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Result r = Result.Succeeded;
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

            //set the formatting options
            ExcelTextFormat format = new ExcelTextFormat();
            format.Delimiter = ';';
            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
            format.Culture.DateTimeFormat.ShortDatePattern = "dd-mm-yyyy";
            format.Encoding = new UTF8Encoding();

            ////Instantiate the Application object.
            //xlApp.Application excelApp = new xlApp.Application();

            ////Open the excel file.
            //workbook = excelApp.Workbooks.Open(@"c:\\temp\\test.xlsx", 0, false, 5, "", "",
            //               false, XlPlatform.xlWindows, "",
            //               true, false, 0, true, false, false)

            //create a new Excel package


            foreach (ViewSchedule vs in col)
            {
                vs.Export(@"c:\\temp\\test", Environment.UserName + vs.Name + ".csv", opt);

                //read the CSV file from disk
                FileInfo file = new FileInfo(@"c:\\temp\\test" + Environment.UserName + vs.Name + ".csv");

                var newFile = (@"c:\\temp\\test" + Environment.UserName + vs.Name + ".csv");

                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    //create a WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(vs.Name);

                    //load the CSV data into cell A1
                    worksheet.Cells["A1"].LoadFromText(file, format);

                    // file name with .xlsx extension  
                    string p_strPath = @"c:\\temp\\test.xlsx";

                    if (File.Exists(p_strPath))
                        File.Delete(p_strPath);

                    // Create excel file on physical disk  
                    FileStream objFileStrm = File.Create(p_strPath);
                    objFileStrm.Close();

                    // Write content to excel file  
                    File.WriteAllBytes(p_strPath, excelPackage.GetAsByteArray());
                    //Close Excel package 
                    excelPackage.Dispose();
                }

                //newSheet = (Worksheet)sheets.Add(sheets[i], Type.Missing, Type.Missing, Type.Missing);

                // workbook.ActiveSheet(@"c:\\temp\\test\" + vs.Name + ".csv");
                //   newSheet.Name = vs.Name;
                //   workbook.Save();
            }

            //workbook.SaveAs(@"c:\\temp\\DIKKETEST", ".xlsx");
            //workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                ExcelApp.Quit();
                Marshal.ReleaseComObject(newSheet);
                Marshal.ReleaseComObject(sheets);
                Marshal.ReleaseComObject(workbook);

                return r;
           
        }
    } 
}
