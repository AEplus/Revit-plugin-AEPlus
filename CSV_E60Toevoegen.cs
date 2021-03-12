using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;

namespace E_60Toevoegen
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class CSV_E60Toevoegen : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            System.IO.Directory.CreateDirectory(@"c:\\temp\\E_60");
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

            // Creates new excelpackage this 
            using (ExcelPackage excelEngine = new ExcelPackage())
            {
                foreach (ViewSchedule vs in col)
                {
                    // Searches for schedules containing AE E60 M52 en M57 ventilatierooster
                    // dit zijn de schedules waarbij het met aantallen is.
                    if (vs.Name.Contains("AE_E60")
                        || vs.Name.Contains("AE_M52")
                        || vs.Name.Contains("AE_M57_ Ventilatieroosters")
                        || vs.Name.Contains("AE_M57_Toestellen VENT")
                        || vs.Name.Contains("AE_M50_Toestellen HVAC coll"))
                    {
                        //create a WorkSheet
                        ExcelWorksheet ws1 = excelEngine.Workbook.Worksheets.Add(vs.Name);
                        // Export c:\\temp --> Will be save as
                        vs.Export(@"c:\\temp\\E_60\", Environment.UserName + vs.Name + ".csv", opt);
                        FileInfo file = new FileInfo(@"c:\\temp\\E_60\" + Environment.UserName + vs.Name + ".csv");

                        // Adds Worksheet as first in the row 
                        ws1.Workbook.Worksheets.MoveToStart(vs.Name);

                        // Formating for writing to xlsx
                        var format = new ExcelTextFormat()
                        {
                            Culture = CultureInfo.InvariantCulture,
                            // Escape character for values containing the Delimiter
                            // ex: "A,Name",1 --> two cells, not three
                            TextQualifier = '"'
                            // Other properties
                            // EOL, DataTypes, Encoding, SkipLinesBeginning/End
                        };

                        ws1.Cells["A1"].LoadFromText(file, format);

                        // excelEngine.Workbook.Worksheets.MoveBefore(i);
                        // the path of the file
                        string filePath = "C:\\temp\\E_60\\Excel_E_60.xlsx";

                        // Write the file to the disk
                        FileInfo fi = new FileInfo(filePath);
                        excelEngine.SaveAs(fi);

                        // Excel terug uitlezen en kijken naar welke cell er leeg in de eerste kolom
                        // om vervolgens de rij te kopieren. 
                        //Create COM Objects. Create a COM object for everything that is referenced
                        Excel.Application xlApp = new Excel.Application();
                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\\temp\\E_60\\Excel_E_60.xlsx");
                        // Excel.Workbook wb2 = xlApp.Workbooks.Open("c:\\temp\\E_60\\Missing.xlsx");

                        foreach (Excel.Worksheet xlworksheet in xlWorkbook.Worksheets)
                        {



                        }
                        //    Microsoft.Office.Interop.Excel.Worksheet x = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                        //    Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        //    Excel.Range xlRange = xlWorksheet.UsedRange;
                        //    int totalRows = xlRange.Rows.Count;
                        //    int totalColumns = xlRange.Columns.Count;
                        //    int i = 0;
                        //    for (i = 1; i <= totalRows; i++)
                        //    {
                        //        // Conditional Needed to check if column row is blank
                        //        if (x.Range["A" + i].Value == "")
                        //        {
                        //            Excel.Range src = xlWorkbook.Sheets[xlworksheet].Range("A" + i + ":G" + i)
                        //            (x.Rows[i] as Microsoft.Office.Interop.Excel.Range).Delete();

                        //        }
                        //        xlWorkbook.Close(true, Type.Missing, Type.Missing);
                        //        xlApp.Quit();
                        //    }


                        //    string colToCheck = "A:A";
                        //    string SearchNo = "";
                        //    Excel.Range resultRange;
                        //    Excel.Range colRange = xlworksheet.Range[colToCheck];
                        //    resultRange = colRange.Find(
                        //                    What: SearchNo,
                        //                    LookIn: Excel.XlFindLookIn.xlValues,
                        //                    LookAt: Excel.XlLookAt.xlPart,
                        //                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                        //                    SearchDirection: Excel.XlSearchDirection.xlNext
                        //    );
                        //    int findrow = resultRange.Row;


                        //    Worksheet srcSheet = xlWorkbook.Worksheets[0];
                        //    Worksheet destSheet = xlWorkbook.Worksheets[1];


                        //    //Source range to be copied
                        //    Excel.Range sourceRange = srcSheet.Cells.CreateRange("A" + i + ":J" + i);

                        //    //Destination range
                        //    Excel.Range destRange = destSheet.Cells.CreateRange();

                        //    //Copy the source range into destination range
                        //    destRange.Copy(sourceRange);

                        //    Excel.Workbook.Save(@"c:\\temp\\E_60\\Missing.xlsx");
                        //}

                    }
                }
                // result succeeded 
                return r;
            }
        }
    }
}



