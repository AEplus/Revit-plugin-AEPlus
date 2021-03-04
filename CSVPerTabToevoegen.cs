using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using xlApp = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
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
            ViewScheduleExportOptions opt = new ViewScheduleExportOptions();
            //{
            //    TextQualifier = ExportTextQualifier.None,
            //    FieldDelimiter = ","
            //};

            

            foreach (ViewSchedule vs in col)
            {
                if (vs.Name.Contains("AE")){
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

                // read the CSV file from disk
                // FileInfo file = new FileInfo(@"c:\\temp\\test" + Environment.UserName + vs.Name + ".csv");

                // var newFile = (@"c:\\temp\\test" + Environment.UserName + vs.Name + ".csv");


                // DIT IS EPPlus package 
                //using (ExcelPackage excelPackage = new ExcelPackage())
                //{
                //    //create a WorkSheet
                //    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(vs.Name);

                //    //load the CSV data into cell A1
                //    worksheet.Cells["A1"].LoadFromText(file, format);

                //    // file name with .xlsx extension  
                //    string p_strPath = @"c:\\temp\\test.xlsx";

                //    if (File.Exists(p_strPath))
                //        File.Delete(p_strPath);

                //    // Create excel file on physical disk  
                //    FileStream objFileStrm = File.Create(p_strPath);
                //    objFileStrm.Close();

                //    // Write content to excel file  
                //    File.WriteAllBytes(p_strPath, excelPackage.GetAsByteArray());
                //    //Close Excel package 
                //    excelPackage.Dispose();
                // TOT HIER IS HET EPPlus Package

            }
            //newSheet = (Worksheet)sheets.Add(sheets[i], Type.Missing, Type.Missing, Type.Missing);

            // workbook.ActiveSheet(@"c:\\temp\\test\" + vs.Name + ".csv");
            //   newSheet.Name = vs.Name;
            //   workbook.Save();
            ExcelApp.Quit();
            Marshal.ReleaseComObject(newSheet);
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(workbook);

            return r;
        }
        //workbook.SaveAs(@"c:\\temp\\DIKKETEST", ".xlsx");
        //workbook.Close(Type.Missing, Type.Missing, Type.Missing);
    }
}